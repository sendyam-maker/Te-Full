VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm Frmacc0000 
   BackColor       =   &H8000000C&
   Caption         =   "財務管理系統"
   ClientHeight    =   6240
   ClientLeft      =   132
   ClientTop       =   780
   ClientWidth     =   9492
   Icon            =   "Frmacc0000.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '系統預設值
   WindowState     =   2  '最大化
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '對齊表單下方
      Height          =   384
      Left            =   0
      TabIndex        =   1
      Top             =   5856
      Width           =   9492
      _ExtentX        =   16743
      _ExtentY        =   677
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1620
      Top             =   330
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9492
      _ExtentX        =   16743
      _ExtentY        =   1016
      ButtonWidth     =   487
      ButtonHeight    =   889
      Appearance      =   1
      _Version        =   393216
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   900
      Top             =   360
   End
   Begin VB.Timer Timer2 
      Left            =   480
      Top             =   780
   End
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   360
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   3000
      _ExtentX        =   593
      _ExtentY        =   593
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
   Begin VB.Menu Main1 
      Caption         =   "國內應收(&I)"
      Begin VB.Menu Main1_1 
         Caption         =   "基本資料"
         Begin VB.Menu Main1_1_1 
            Caption         =   "收據"
            Index           =   1
            Begin VB.Menu Main1_1_1_1 
               Caption         =   "收據開立作業-整批"
               Index           =   1
            End
            Begin VB.Menu Main1_1_1_1 
               Caption         =   "收據開立作業"
               Index           =   2
            End
            Begin VB.Menu Main1_1_1_1 
               Caption         =   "收據/請款單作廢作業"
               Index           =   3
            End
            Begin VB.Menu Main1_1_1_1 
               Caption         =   "收據抬頭修改"
               Index           =   4
            End
            Begin VB.Menu Main1_1_1_1 
               Caption         =   "收據金額修改"
               Index           =   5
            End
            Begin VB.Menu Main1_1_1_1 
               Caption         =   "拆收據作業"
               Index           =   6
            End
            Begin VB.Menu Main1_1_1_1 
               Caption         =   "國內收據產生INVOICE"
               Index           =   7
            End
            Begin VB.Menu Main1_1_1_1 
               Caption         =   "INVOICE編號作廢作業"
               Index           =   8
            End
         End
         Begin VB.Menu Main1_1_1 
            Caption         =   "收款/銷退/暫收"
            Index           =   2
            Begin VB.Menu Main1_1_1_2 
               Caption         =   "收款作業-整批"
               Index           =   1
            End
            Begin VB.Menu Main1_1_1_2 
               Caption         =   "收款作業"
               Index           =   2
            End
            Begin VB.Menu Main1_1_1_2 
               Caption         =   "銷帳退費作業"
               Index           =   3
            End
            Begin VB.Menu Main1_1_1_2 
               Caption         =   "暫收款作業"
               Index           =   4
            End
         End
         Begin VB.Menu Main1_1_1 
            Caption         =   "發票"
            Index           =   3
            Begin VB.Menu Main1_1_1_3 
               Caption         =   "請款單開立發票作業"
               Index           =   1
            End
            Begin VB.Menu Main1_1_1_3 
               Caption         =   "發票作廢作業"
               Index           =   2
            End
            Begin VB.Menu Main1_1_1_3 
               Caption         =   "發票申報作業"
               Index           =   3
            End
            Begin VB.Menu Main1_1_1_3 
               Caption         =   "發票跨期轉開作業"
               Index           =   4
            End
            Begin VB.Menu Main1_1_1_3 
               Caption         =   "發票上傳作業"
               Index           =   5
            End
            Begin VB.Menu Main1_1_1_3 
               Caption         =   "發票號碼維護"
               Index           =   6
            End
            Begin VB.Menu Main1_1_1_3 
               Caption         =   "特殊發票客戶資料維護及查詢"
               Index           =   7
            End
            Begin VB.Menu Main1_1_1_3 
               Caption         =   "發票備註維護作業"
               Index           =   8
            End
         End
         Begin VB.Menu Main1_1_1 
            Caption         =   "其他"
            Index           =   4
            Begin VB.Menu Main1_1_1_4 
               Caption         =   "國內應收待處理作業"
               Index           =   1
            End
            Begin VB.Menu Main1_1_1_4 
               Caption         =   "收據抬頭基本資料維護"
               Index           =   2
            End
            Begin VB.Menu Main1_1_1_4 
               Caption         =   "退費收訖憑單維護"
               Index           =   3
            End
            Begin VB.Menu Main1_1_1_4 
               Caption         =   "手開收據開立"
               Index           =   4
            End
            Begin VB.Menu Main1_1_1_4 
               Caption         =   "收文金額分配作業"
               Index           =   5
            End
            Begin VB.Menu Main1_1_1_4 
               Caption         =   "客戶電匯資料維護及查詢"
               Index           =   6
            End
            Begin VB.Menu Main1_1_1_4 
               Caption         =   "客戶應收帳款收文檢查上限"
               Index           =   7
            End
            Begin VB.Menu Main1_1_1_4 
               Caption         =   "客戶特殊付款週期維護"
               Index           =   8
            End
         End
      End
      Begin VB.Menu Line1_1 
         Caption         =   "-"
      End
      Begin VB.Menu Main1_2 
         Caption         =   "查詢作業"
         Begin VB.Menu Main1_2_1 
            Caption         =   "收據/請款單,發票資料查詢"
         End
         Begin VB.Menu Main1_2_14 
            Caption         =   "應收帳款綜合查詢"
         End
         Begin VB.Menu Main1_2_2 
            Caption         =   "客戶帳款查詢"
         End
         Begin VB.Menu Main1_2_3 
            Caption         =   "智權人員帳款查詢"
         End
         Begin VB.Menu Main1_2_4 
            Caption         =   "本所案號帳目查詢"
         End
         Begin VB.Menu Main1_2_5 
            Caption         =   "收款單號查詢"
         End
         Begin VB.Menu Main1_2_6 
            Caption         =   "收據/請款單作廢查詢"
         End
         Begin VB.Menu Main1_2_7 
            Caption         =   "收文與收據資料檢核查詢"
         End
         Begin VB.Menu Main1_2_8 
            Caption         =   "手開收據資料查詢"
         End
         Begin VB.Menu Main1_2_9 
            Caption         =   "客戶應收帳款查詢"
         End
         Begin VB.Menu Main1_2_10 
            Caption         =   "未列印收據/請款單查詢"
         End
         Begin VB.Menu Main1_2_11 
            Caption         =   "發票資料查詢"
         End
         Begin VB.Menu Main1_2_12 
            Caption         =   "已開發票未收款明細查詢"
         End
         Begin VB.Menu Main1_2_13 
            Caption         =   "價目表查詢"
         End
         Begin VB.Menu Main1_2_15 
            Caption         =   "法律與智慧所案件對照表"
         End
         Begin VB.Menu Main1_2_16 
            Caption         =   "收文金額異常檢查"
         End
      End
      Begin VB.Menu Line1_2 
         Caption         =   "-"
      End
      Begin VB.Menu Main1_4 
         Caption         =   "報表列印"
         Begin VB.Menu Main1_4_1 
            Caption         =   "收據列印"
            Index           =   1
         End
         Begin VB.Menu Main1_4_1 
            Caption         =   "補開收據列印"
            Index           =   2
         End
         Begin VB.Menu Main1_4_1 
            Caption         =   "客戶對帳單"
            Index           =   3
         End
         Begin VB.Menu Main1_4_1 
            Caption         =   "客戶帳款明細表"
            Index           =   4
         End
         Begin VB.Menu Main1_4_1 
            Caption         =   "智權人員帳款明細表"
            Index           =   5
         End
         Begin VB.Menu Main1_4_1 
            Caption         =   "智權人員應收規費明細表"
            Index           =   6
         End
         Begin VB.Menu Main1_4_1 
            Caption         =   "國內帳齡分析表"
            Index           =   7
         End
         Begin VB.Menu Main1_4_1 
            Caption         =   "國內人員收文點數及收款點數統計"
            Index           =   8
         End
         Begin VB.Menu Main1_4_1 
            Caption         =   "銷帳退費明細表"
            Index           =   9
         End
         Begin VB.Menu Main1_4_1 
            Caption         =   "暫收款明細表"
            Index           =   10
         End
         Begin VB.Menu Main1_4_1 
            Caption         =   "收據/請款單作廢明細表"
            Index           =   11
         End
         Begin VB.Menu Main1_4_1 
            Caption         =   "收文與收據資料檢核表"
            Index           =   12
         End
         Begin VB.Menu Main1_4_1 
            Caption         =   "收據抬頭修改清單"
            Index           =   13
         End
         Begin VB.Menu Main1_4_1 
            Caption         =   "收據帳款明細列印"
            Index           =   14
         End
         Begin VB.Menu Main1_4_1 
            Caption         =   "收款扣繳改年度清單"
            Index           =   15
         End
         Begin VB.Menu Main1_4_1 
            Caption         =   "請款單列印"
            Index           =   16
         End
         Begin VB.Menu Main1_4_1 
            Caption         =   "補開請款單列印"
            Index           =   17
         End
         Begin VB.Menu Main1_4_1 
            Caption         =   "同仁介紹案源獎金明細表"
            Index           =   18
         End
         Begin VB.Menu Main1_4_1 
            Caption         =   "銷貨退回折讓單列印"
            Index           =   19
            Visible         =   0   'False
         End
         Begin VB.Menu Main1_4_1 
            Caption         =   "國內收據產生特殊請款單"
            Index           =   20
         End
         Begin VB.Menu Main1_4_1 
            Caption         =   "應收帳款財務處控管資料表"
            Index           =   21
         End
         Begin VB.Menu Main1_4_1 
            Caption         =   "智權人員請款明細表"
            Index           =   22
         End
      End
   End
   Begin VB.Menu Main8 
      Caption         =   "國內應付(&O)"
      Begin VB.Menu Main8_1 
         Caption         =   "基本資料"
         Begin VB.Menu Main8_1_1 
            Caption         =   "廠商基本資料"
            Index           =   1
         End
         Begin VB.Menu Main8_1_1 
            Caption         =   "應付款資料"
            Index           =   2
         End
         Begin VB.Menu Main8_1_1 
            Caption         =   "付款作業"
            Index           =   3
         End
         Begin VB.Menu Main8_1_1 
            Caption         =   "員工翻譯費率維護"
            Index           =   4
         End
         Begin VB.Menu Main8_1_1 
            Caption         =   "翻譯費資料輸入"
            Index           =   5
         End
         Begin VB.Menu Main8_1_1_1 
            Caption         =   "廠商 / 客戶 / 員工"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu Main8_1_1_1 
            Caption         =   "台一 / 智權"
            Index           =   2
            Visible         =   0   'False
         End
      End
      Begin VB.Menu Line8_0 
         Caption         =   "-"
      End
      Begin VB.Menu Main8_2 
         Caption         =   "查詢作業"
         Begin VB.Menu Main8_2_1 
            Caption         =   "應付款查詢"
            Index           =   1
         End
         Begin VB.Menu Main8_2_1 
            Caption         =   "客戶回執查詢"
            Index           =   2
         End
         Begin VB.Menu Main8_2_1 
            Caption         =   "翻譯費查詢"
            Index           =   3
         End
         Begin VB.Menu Main8_2_1 
            Caption         =   "翻譯完稿案件查詢/列印"
            Index           =   4
         End
         Begin VB.Menu Main8_2_1 
            Caption         =   "翻譯費用及請款明細查詢/列印"
            Index           =   5
         End
         Begin VB.Menu Main8_2_1 
            Caption         =   "出庭費查詢"
            Index           =   6
         End
      End
      Begin VB.Menu Line8_1 
         Caption         =   "-"
      End
      Begin VB.Menu Main8_4 
         Caption         =   "報表列印"
         Begin VB.Menu Main8_4_1 
            Caption         =   "付款工作底稿"
            Index           =   1
         End
         Begin VB.Menu Main8_4_1 
            Caption         =   "客戶付款明細"
            Index           =   2
         End
         Begin VB.Menu Main8_4_1 
            Caption         =   "國內廠商付款明細表"
            Index           =   3
         End
         Begin VB.Menu Main8_4_1 
            Caption         =   "應付款統計表"
            Index           =   4
         End
         Begin VB.Menu Main8_4_1 
            Caption         =   "名條列印"
            Index           =   5
         End
         Begin VB.Menu Main8_4_1 
            Caption         =   "廠商付款明細表"
            Index           =   6
         End
         Begin VB.Menu Main8_4_1 
            Caption         =   "翻譯費總表"
            Index           =   7
         End
         Begin VB.Menu Main8_4_1 
            Caption         =   "翻譯費明細表"
            Index           =   8
         End
      End
   End
   Begin VB.Menu Main2 
      Caption         =   "國外應收/付(&F)"
      Begin VB.Menu Main2_1 
         Caption         =   "基本資料"
         Begin VB.Menu Main2_1_1 
            Caption         =   "收款作業"
         End
         Begin VB.Menu Main2_1_2 
            Caption         =   "暫收款作業"
         End
         Begin VB.Menu Main2_1_3 
            Caption         =   "暫收款退費作業"
         End
         Begin VB.Menu Main2_1_4 
            Caption         =   "銷帳作業"
         End
         Begin VB.Menu Main2_1_18 
            Caption         =   "國外收款分析表"
         End
         Begin VB.Menu Main2_1_7 
            Caption         =   "結匯資料輸入"
         End
         Begin VB.Menu Main2_1_9 
            Caption         =   "調整付款明細"
         End
         Begin VB.Menu Main2_1_10 
            Caption         =   "結匯匯率輸入"
         End
         Begin VB.Menu Main2_1_11 
            Caption         =   "匯票輸入"
         End
         Begin VB.Menu Main2_1_12 
            Caption         =   "付款後退費作業"
         End
         Begin VB.Menu Main2_1_13 
            Caption         =   "抵帳作業"
         End
         Begin VB.Menu Main2_1_19 
            Caption         =   "銀存匯率資料輸入"
         End
         Begin VB.Menu Main2_1_20 
            Caption         =   "客戶/代理人匯款銀行資料維護"
         End
         Begin VB.Menu Main2_1_21 
            Caption         =   "客戶/代理人財務EMail資料維護"
         End
         Begin VB.Menu Main2_1_22 
            Caption         =   "國外固定寄催款單代理人維護"
         End
         Begin VB.Menu Main6_1_5 
            Caption         =   "帳單輸入"
         End
         Begin VB.Menu Main6_1_11 
            Caption         =   "主管審核作業"
         End
         Begin VB.Menu Main6_1_7 
            Caption         =   "抵帳單輸入"
         End
         Begin VB.Menu Main6_1_6 
            Caption         =   "帳單作廢作業"
         End
         Begin VB.Menu Main6_1_2 
            Caption         =   "請款單輸入"
         End
         Begin VB.Menu Main6_1_3 
            Caption         =   "折讓輸入"
         End
         Begin VB.Menu Main6_1_4 
            Caption         =   "請款單作廢作業"
         End
         Begin VB.Menu Main6_1_8 
            Caption         =   "FC案件不請款確認維護"
         End
         Begin VB.Menu Main6_1_9 
            Caption         =   "預估結匯匯率資料維護"
         End
         Begin VB.Menu Main6_1_12 
            Caption         =   "美金請款匯率資料維護"
         End
         Begin VB.Menu Main6_1_13 
            Caption         =   "其他幣別請款匯率資料維護"
         End
         Begin VB.Menu Main6_1_14 
            Caption         =   "國外付款日期調整"
         End
         Begin VB.Menu Main6_1_10 
            Caption         =   "電子結匯作業"
         End
      End
      Begin VB.Menu Line2_1 
         Caption         =   "-"
      End
      Begin VB.Menu Main2_2 
         Caption         =   "查詢作業"
         Begin VB.Menu Main2_2_1 
            Caption         =   "國外代理人帳目查詢"
         End
         Begin VB.Menu Main2_2_2 
            Caption         =   "國外案件帳目查詢"
         End
         Begin VB.Menu Main2_2_3 
            Caption         =   "國外請款金額查詢"
         End
         Begin VB.Menu Main2_2_4 
            Caption         =   "其他結匯查詢"
         End
         Begin VB.Menu Main2_2_5 
            Caption         =   "未請款查詢"
         End
         Begin VB.Menu Main2_2_6 
            Caption         =   "FC收款請款點數查詢"
         End
         Begin VB.Menu Main2_2_7 
            Caption         =   "國外部智權人員帳款查詢"
         End
         Begin VB.Menu Main2_2_8 
            Caption         =   "各幣別最新請款匯率查詢"
         End
         Begin VB.Menu Main2_2_9 
            Caption         =   "FC業務請款／收款明細表(國外部用)"
         End
         Begin VB.Menu Main2_2_10 
            Caption         =   "FC專業收款明細表(國外部用)"
         End
      End
      Begin VB.Menu Line2_2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu Main2_3 
         Caption         =   "批次作業"
         Visible         =   0   'False
         Begin VB.Menu Main2_3_1 
            Caption         =   "請款明細刪除作業"
         End
      End
      Begin VB.Menu Line2_3 
         Caption         =   "-"
      End
      Begin VB.Menu Main2_4 
         Caption         =   "報表列印"
         Begin VB.Menu Main2_4_6 
            Caption         =   "FC收據"
         End
         Begin VB.Menu Main2_4_7 
            Caption         =   "FC催款單"
         End
         Begin VB.Menu Main2_4_8 
            Caption         =   "FC請款單"
         End
         Begin VB.Menu Main2_4_9 
            Caption         =   "國外帳款對帳單"
         End
         Begin VB.Menu Main2_4_10 
            Caption         =   "國外帳齡分析表"
         End
         Begin VB.Menu Main2_4_13 
            Caption         =   "代理人FC帳款明細表"
         End
         Begin VB.Menu Main2_4_14 
            Caption         =   "國外應收規費、服務費分析表"
         End
         Begin VB.Menu Main2_4_3 
            Caption         =   "國外結匯水單列印"
         End
         Begin VB.Menu Main2_4_20 
            Caption         =   "國外結匯媒體檔產生作業"
         End
         Begin VB.Menu Main2_4_4 
            Caption         =   "結匯明細表(傳票附件,暫不使用)"
         End
         Begin VB.Menu Main2_4_5 
            Caption         =   "國外付款明細表"
         End
         Begin VB.Menu Main6_4_1 
            Caption         =   "代理人帳目排名"
         End
         Begin VB.Menu Main6_4_3 
            Caption         =   "代理人逾期帳款分析表"
         End
         Begin VB.Menu Main2_4_15 
            Caption         =   "帳單輸入三個月未結匯明細"
         End
         Begin VB.Menu Main2_4_16 
            Caption         =   "地址條列印"
         End
         Begin VB.Menu Main2_4_18 
            Caption         =   "國外部FC案件不請款清單"
         End
         Begin VB.Menu Main2_4_19 
            Caption         =   "國外請款點數分析表"
         End
         Begin VB.Menu Main2_4_21 
            Caption         =   "國外請款單產生國內收據"
         End
      End
   End
   Begin VB.Menu Main3 
      Caption         =   "票據(&C)"
      Begin VB.Menu Main3_1 
         Caption         =   "基本資料"
         Begin VB.Menu Main3_1_1 
            Caption         =   "收票作業"
         End
         Begin VB.Menu Main3_1_2 
            Caption         =   "開票作業"
         End
         Begin VB.Menu Main3_1_12 
            Caption         =   "票據轉出作業"
         End
         Begin VB.Menu Main3_1_3 
            Caption         =   "票據託收作業"
         End
         Begin VB.Menu Main3_1_11 
            Caption         =   "即期票存入作業"
         End
         Begin VB.Menu Main3_1_8 
            Caption         =   "銀行基本資料"
         End
         Begin VB.Menu Main3_1_9 
            Caption         =   "銀行帳戶基本資料"
         End
         Begin VB.Menu Main3_1_4 
            Caption         =   "支票未領備註說明"
         End
         Begin VB.Menu Main3_1_5 
            Caption         =   "退票作業"
         End
         Begin VB.Menu Main3_1_6 
            Caption         =   "抽票作業"
         End
         Begin VB.Menu Main3_1_10 
            Caption         =   "票據作廢作業"
         End
         Begin VB.Menu Main3_1_13 
            Caption         =   "銀行往來調節作業"
         End
         Begin VB.Menu Main3_1_7 
            Caption         =   "票據貼現作業"
         End
         Begin VB.Menu Main3_1_14 
            Caption         =   "智慧局送件付款作業"
         End
         Begin VB.Menu Main3_1_15 
            Caption         =   "分所智慧局送件付款作業"
         End
         Begin VB.Menu Main3_1_16 
            Caption         =   "智慧局電子送件付款作業"
         End
         Begin VB.Menu Main3_1_17 
            Caption         =   "智慧局電子送件網路扣帳作業"
         End
      End
      Begin VB.Menu Line3_1 
         Caption         =   "-"
      End
      Begin VB.Menu Main3_2 
         Caption         =   "查詢作業"
         Begin VB.Menu Main3_2_1 
            Caption         =   "銀行帳號流動資金查詢"
         End
         Begin VB.Menu Main3_2_2 
            Caption         =   "銀行帳號別票據彙總查詢"
            Visible         =   0   'False
         End
         Begin VB.Menu Main3_2_3 
            Caption         =   "銀行帳號別票據明細查詢"
            Visible         =   0   'False
         End
         Begin VB.Menu Main3_2_5 
            Caption         =   "到期日別票據明細查詢"
         End
         Begin VB.Menu Main3_2_15 
            Caption         =   "兌現日別票據明細查詢"
         End
         Begin VB.Menu Main3_2_6 
            Caption         =   "收票資料查詢"
         End
         Begin VB.Menu Main3_2_7 
            Caption         =   "開票資料查詢"
         End
         Begin VB.Menu Main3_2_8 
            Caption         =   "託收票據資料查詢"
         End
         Begin VB.Menu Main3_2_10 
            Caption         =   "退票資料查詢"
         End
         Begin VB.Menu Main3_2_11 
            Caption         =   "抽票資料查詢"
         End
         Begin VB.Menu Main3_2_12 
            Caption         =   "貼現票據資料查詢"
         End
         Begin VB.Menu Main3_2_13 
            Caption         =   "甲存支兌領作業查詢"
         End
         Begin VB.Menu Main3_2_14 
            Caption         =   "銀行往來調節查詢"
         End
      End
      Begin VB.Menu Line3_2 
         Caption         =   "-"
      End
      Begin VB.Menu Main3_3 
         Caption         =   "批次作業"
         Begin VB.Menu Main3_3_1 
            Caption         =   "票據兌現處理"
         End
         Begin VB.Menu Main3_3_2 
            Caption         =   "票據歷史資料刪除處理"
         End
      End
      Begin VB.Menu Line3_3 
         Caption         =   "-"
      End
      Begin VB.Menu Main3_4 
         Caption         =   "報表列印"
         Begin VB.Menu Main3_4_19 
            Caption         =   "支票列印"
         End
         Begin VB.Menu Main3_4_1 
            Caption         =   "應收票據資料表"
         End
         Begin VB.Menu Main3_4_2 
            Caption         =   "應付票據資料表"
         End
         Begin VB.Menu Main3_4_3 
            Caption         =   "託收票據資料表"
         End
         Begin VB.Menu Main3_4_4 
            Caption         =   "銀行帳號別票據彙總表"
            Visible         =   0   'False
         End
         Begin VB.Menu Main3_4_5 
            Caption         =   "銀行帳號別票據明細表"
            Visible         =   0   'False
         End
         Begin VB.Menu Main3_4_16 
            Caption         =   "兌現日別資金流動彙總表"
         End
         Begin VB.Menu Main3_4_6 
            Caption         =   "兌現日別票據明細表"
         End
         Begin VB.Menu Main3_4_7 
            Caption         =   "往來對象別票據彙總表"
         End
         Begin VB.Menu Main3_4_8 
            Caption         =   "往來對象別票據明細表"
         End
         Begin VB.Menu Main3_4_9 
            Caption         =   "退票資料表"
         End
         Begin VB.Menu Main3_4_10 
            Caption         =   "抽票資料表"
         End
         Begin VB.Menu Main3_4_11 
            Caption         =   "票據貼現資料檢核表"
         End
         Begin VB.Menu Main3_4_12 
            Caption         =   "銀行帳號別資金流動表"
         End
         Begin VB.Menu Main3_4_13 
            Caption         =   "日期別資金流動預測表"
         End
         Begin VB.Menu Main3_4_14 
            Caption         =   "銀行調節資料表"
         End
         Begin VB.Menu Main3_4_17 
            Caption         =   "銀行別資料表"
         End
         Begin VB.Menu Main3_4_18 
            Caption         =   "甲存支票未兌領明細表"
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu Main4 
      Caption         =   "總帳(&A)"
      Begin VB.Menu Main4_1 
         Caption         =   "基本資料"
         Begin VB.Menu Main4_1_1 
            Caption         =   "會計科目基本資料"
            Index           =   1
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "傳票輸入"
            Index           =   2
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "公司基本資料"
            Index           =   3
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "部門基本資料"
            Index           =   4
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "預算資料"
            Index           =   5
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "每月固定傳票資料"
            Index           =   6
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "分攤類別資料"
            Index           =   7
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "分攤類別比率資料"
            Index           =   8
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "CF案件結餘結算作業"
            Index           =   9
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "CF案件結餘作廢作業"
            Index           =   10
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "傳票過帳後摘要修改"
            Index           =   12
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "應收付分錄調整"
            Index           =   13
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "簽收作業"
            Index           =   14
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "結餘保留放出產生傳票"
            Index           =   15
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "每月業績開放/關閉輸入"
            Index           =   16
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "智權點數實績與結餘輸入"
            Index           =   17
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "ACS 待分潤"
            Index           =   18
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "非智權結餘轉撥報出傳票產生(隱藏版)"
            Index           =   19
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "智權期末實績保留傳票產生"
            Index           =   20
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "智權期末結餘保留傳票產生"
            Index           =   21
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "智權期末結餘保留資料刪除"
            Index           =   22
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "每月固定傳票資料-非分攤"
            Index           =   23
         End
      End
      Begin VB.Menu Line4_1 
         Caption         =   "-"
      End
      Begin VB.Menu Main4_2 
         Caption         =   "查詢作業"
         Begin VB.Menu Main4_2_1 
            Caption         =   "傳票資料查詢"
            Index           =   1
         End
         Begin VB.Menu Main4_2_1 
            Caption         =   "日記帳查詢"
            Index           =   2
         End
         Begin VB.Menu Main4_2_1 
            Caption         =   "科目分類帳查詢"
            Index           =   3
         End
         Begin VB.Menu Main4_2_1 
            Caption         =   "科目餘額查詢"
            Index           =   4
         End
         Begin VB.Menu Main4_2_1 
            Caption         =   "科目明細查詢(對沖)"
            Index           =   5
         End
         Begin VB.Menu Main4_2_1 
            Caption         =   "智權人員點數查詢"
            Index           =   6
         End
         Begin VB.Menu Main4_2_1 
            Caption         =   "單據/傳票號碼查詢"
            Index           =   7
         End
         Begin VB.Menu Main4_2_1 
            Caption         =   "智權人員結餘點數查詢"
            Index           =   8
         End
         Begin VB.Menu Main4_2_1 
            Caption         =   "簽收資料查詢"
            Index           =   9
         End
         Begin VB.Menu Main4_2_1 
            Caption         =   "業績點數查詢"
            Index           =   10
         End
         Begin VB.Menu Main4_2_1 
            Caption         =   "CF 結餘單查詢(已有結餘單號)"
            Index           =   11
         End
         Begin VB.Menu Main4_2_1 
            Caption         =   "CF 結餘單案件明細查詢"
            Index           =   12
         End
         Begin VB.Menu Main4_2_1 
            Caption         =   "CF 可結餘日期查詢(尚未產生結餘單號)"
            Index           =   13
         End
         Begin VB.Menu Main4_2_1 
            Caption         =   "專業點數分析"
            Index           =   14
         End
         Begin VB.Menu Main4_2_1 
            Caption         =   "未繳款資料查詢與銀存核對"
            Index           =   15
         End
         Begin VB.Menu Main4_2_1 
            Caption         =   "過帳前綜合損益查詢及列印"
            Index           =   16
         End
         Begin VB.Menu Main4_2_1 
            Caption         =   "旅遊補助付款通知"
            Index           =   17
         End
         Begin VB.Menu Main4_2_1 
            Caption         =   "出庭費發放通知"
            Index           =   18
         End
      End
      Begin VB.Menu Line4_2 
         Caption         =   "-"
      End
      Begin VB.Menu Main4_3 
         Caption         =   "批次作業"
         Begin VB.Menu Main4_3_1 
            Caption         =   "應收/付轉傳票作業"
            Index           =   1
         End
         Begin VB.Menu Main4_3_1 
            Caption         =   "過帳及分攤作業"
            Index           =   2
         End
         Begin VB.Menu Main4_3_1 
            Caption         =   "月結算作業"
            Index           =   3
         End
         Begin VB.Menu Main4_3_1 
            Caption         =   "年度結轉作業"
            Index           =   4
         End
         Begin VB.Menu Main4_3_1 
            Caption         =   "歷史傳票資料刪除處理"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu Main4_3_1 
            Caption         =   "傳票轉外帳作業"
            Index           =   6
         End
         Begin VB.Menu Main4_3_1 
            Caption         =   "翻譯費轉應付作業"
            Index           =   7
         End
         Begin VB.Menu Main4_3_1 
            Caption         =   "已開發票未收款沖帳作業"
            Index           =   8
         End
         Begin VB.Menu Main4_3_1 
            Caption         =   "取消過帳或月(年)結"
            Index           =   10
         End
      End
      Begin VB.Menu Line4_4 
         Caption         =   "-"
      End
      Begin VB.Menu Main4_5 
         Caption         =   "財產目錄"
         Begin VB.Menu Main4_5_1 
            Caption         =   "財產目錄作業"
            Index           =   1
         End
         Begin VB.Menu Main4_5_1 
            Caption         =   "財產報廢作業"
            Index           =   2
         End
         Begin VB.Menu Main4_5_1 
            Caption         =   "財產目錄表"
            Index           =   3
         End
      End
      Begin VB.Menu Line4_3 
         Caption         =   "-"
      End
      Begin VB.Menu Main4_4 
         Caption         =   "報表列印"
         Begin VB.Menu Main4_4_1 
            Caption         =   "會計傳票列印"
            Index           =   1
         End
         Begin VB.Menu Main4_4_1 
            Caption         =   "日計表"
            Index           =   2
         End
         Begin VB.Menu Main4_4_1 
            Caption         =   "科目明細表(對沖)"
            Index           =   3
         End
         Begin VB.Menu Main4_4_1 
            Caption         =   "科目餘額表"
            Index           =   4
         End
         Begin VB.Menu Main4_4_1 
            Caption         =   "試算表"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu Main4_4_1 
            Caption         =   "科目分類帳"
            Index           =   6
         End
         Begin VB.Menu Main4_4_1 
            Caption         =   "智權人員點數明細表"
            Index           =   7
         End
         Begin VB.Menu Main4_4_1 
            Caption         =   "專業點數明細表"
            Index           =   8
         End
         Begin VB.Menu Main4_4_1 
            Caption         =   "ACS待分潤明細表"
            Index           =   9
         End
         Begin VB.Menu Main4_4_1 
            Caption         =   "專業達成點數表-秘書"
            Index           =   10
         End
         Begin VB.Menu Main4_4_1 
            Caption         =   "預算實績比較表"
            Index           =   11
         End
         Begin VB.Menu Main4_4_1 
            Caption         =   "綜合損益比較表"
            Index           =   12
         End
         Begin VB.Menu Main4_4_1 
            Caption         =   "年度綜合損益統計表"
            Index           =   13
         End
         Begin VB.Menu Main4_4_1 
            Caption         =   "資產負債表"
            Index           =   14
         End
         Begin VB.Menu Main4_4_1 
            Caption         =   "綜合損益表"
            Index           =   15
         End
         Begin VB.Menu Main4_4_1 
            Caption         =   "年度部門綜合損益統計表"
            Index           =   16
            Visible         =   0   'False
         End
         Begin VB.Menu Main4_4_1 
            Caption         =   "資產負債比較表"
            Index           =   17
         End
         Begin VB.Menu Main4_4_1 
            Caption         =   "部門費用統計表"
            Index           =   18
            Visible         =   0   'False
         End
         Begin VB.Menu Main4_4_1 
            Caption         =   "部門綜合損益表(子科目)"
            Index           =   19
            Visible         =   0   'False
         End
         Begin VB.Menu Main4_4_1 
            Caption         =   "部門綜合損益表"
            Index           =   20
         End
         Begin VB.Menu Main4_4_1 
            Caption         =   "預算資料表"
            Index           =   21
         End
         Begin VB.Menu Main4_4_1 
            Caption         =   "會計科目代號對照表"
            Index           =   22
         End
         Begin VB.Menu Main4_4_1 
            Caption         =   "費用科目分攤比率表"
            Index           =   23
         End
         Begin VB.Menu Main4_4_1 
            Caption         =   "結餘單列印"
            Index           =   24
         End
      End
   End
   Begin VB.Menu Main9 
      Caption         =   "扣繳作業(&W)"
      Begin VB.Menu Main9_1 
         Caption         =   "基本資料"
         Begin VB.Menu Main9_1_1 
            Caption         =   "補扣繳作業"
         End
         Begin VB.Menu Main9_1_2 
            Caption         =   "扣繳憑單維護"
         End
         Begin VB.Menu Main9_1_3 
            Caption         =   "扣單作廢作業"
         End
         Begin VB.Menu Main9_1_4 
            Caption         =   "扣繳稅款沖轉作業"
         End
         Begin VB.Menu Main9_1_5 
            Caption         =   "扣繳憑單修正作業"
         End
      End
      Begin VB.Menu Line9_1 
         Caption         =   "-"
      End
      Begin VB.Menu Main9_4 
         Caption         =   "報表查詢及列印"
         Begin VB.Menu Main9_4_1 
            Caption         =   "智權人員別客戶扣繳稅款明細表"
         End
         Begin VB.Menu Main9_4_2 
            Caption         =   "扣繳憑單催收表"
            Visible         =   0   'False
         End
         Begin VB.Menu Main9_4_3 
            Caption         =   "客戶扣繳明細核對表"
         End
         Begin VB.Menu Main9_4_4 
            Caption         =   "扣繳憑單明細表"
         End
         Begin VB.Menu Main9_4_5 
            Caption         =   "繳款書寄出明細"
         End
         Begin VB.Menu Main9_4_7 
            Caption         =   "代填繳款書客戶明細"
         End
         Begin VB.Menu Main9_4_8 
            Caption         =   "補扣繳地址條及清單"
         End
         Begin VB.Menu Main9_4_9 
            Caption         =   "會計師客戶資料查詢"
         End
         Begin VB.Menu Main9_4_10 
            Caption         =   "年度扣繳檢核(抬頭及信箱)"
         End
         Begin VB.Menu Main9_4_11 
            Caption         =   "客戶年度未扣繳查詢"
         End
         Begin VB.Menu Main9_4_6 
            Caption         =   "扣繳憑單查詢及列印"
         End
      End
   End
   Begin VB.Menu Main7 
      Caption         =   "系統(&S)"
      Begin VB.Menu Main7_0 
         Caption         =   "切換連線"
         Visible         =   0   'False
      End
      Begin VB.Menu Main7_1 
         Caption         =   "系統管理"
         Begin VB.Menu Main7_1_1 
            Caption         =   "系統參數變更"
         End
         Begin VB.Menu Main7_1_2 
            Caption         =   "會計科目餘額維護"
         End
         Begin VB.Menu Main7_1_3 
            Caption         =   "報表紙張格式設定"
         End
         Begin VB.Menu Main7_1_4 
            Caption         =   "系統印表機設定"
         End
         Begin VB.Menu Main7_1_5 
            Caption         =   "解除畫面擷取限制"
         End
      End
      Begin VB.Menu Line7_1 
         Caption         =   "-"
      End
      Begin VB.Menu Main7_2 
         Caption         =   "說明"
         Begin VB.Menu Main7_2_2 
            Caption         =   "關於"
         End
      End
      Begin VB.Menu Line7_2 
         Caption         =   "-"
      End
      Begin VB.Menu Main7_3 
         Caption         =   "結束"
      End
   End
End
Attribute VB_Name = "Frmacc0000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2022/2/9 Form2.0不用改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/28 日期欄已修改
Option Explicit

Private Declare Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
Private Declare Function getUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Dim intCounter As Integer
Public m_blnABSActivated As Boolean 'Add By Sindy 2011/9/15
Public m_ChkIsOpenFrm180203 As Boolean 'Add By Sindy 2013/7/8


'Modified by Lydia 2023/11/13 調整國內應收之基本資料的選單
'*****國內應收->基本資料->收據*****
Private Sub Main1_1_1_1_Click(Index As Integer)
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   Select Case Index
      Case 1 '收據開立作業-整批
         If CheckUse("Frmacc1123", strExec) = False Then
            Exit Sub
         End If
         If PUB_GetLock("Frmacc1123", "", "收據開立作業-批次") = False Then
            Exit Sub
         End If
         
         'Added by Morgan 2024/5/24 +避免和單筆同時操作，增加鎖住frmacc1120
         If PUB_GetLock("Frmacc1120", "", "收據開立作業") = False Then
            Call PUB_GetLock("", "Frmacc1123")
            Exit Sub
         End If
         'end 2024/5/24
         
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1123.Show
         Me.MousePointer = vbDefault
            
      Case 2 '收據開立作業
         If CheckUse("Frmacc1120", strExec) = False Then
            Exit Sub
         End If
         If PUB_GetLock("Frmacc1120", "", "收據開立作業") = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1120.Show
         Me.MousePointer = vbDefault
         
      Case 3 '收據/請款單作廢作業
         If CheckUse("Frmacc1130", strExec) = False Then
            Exit Sub
         End If
         tool14_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1130.Show
         Me.MousePointer = vbDefault
         
      Case 4 '收據抬頭修改
         If CheckUse("Frmacc1140", strExec) = False Then
            Exit Sub
         End If
         tool8_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1140.Show
         Me.MousePointer = vbDefault
         
      Case 5 '收據金額修改
         If CheckUse("Frmacc11d0", strExec) = False Then
            Exit Sub
         End If
         tool14_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc11d0.Show
         Me.MousePointer = vbDefault
               
      Case 6 '拆收據作業
         If CheckUse("Frmacc11m0", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc11m0.Show
         Me.MousePointer = vbDefault
      Case 7 '國內收據產生INVOICE
         If CheckUse("Frmacc14o0", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc14o0.Show
         Me.MousePointer = vbDefault
         
      Case 8  'INVOICE編號作廢作業
         If CheckUse("Frmacc14o1", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         MenuDisabled
         Frmacc14o1.Show
   End Select
   
End Sub

'Modified by Lydia 2023/11/13 調整國內應收之基本資料的選單
'*****國內應收->基本資料->收款*****
Private Sub Main1_1_1_2_Click(Index As Integer)
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   
   Select Case Index
      Case 1 '收款作業-整批
         If CheckUse("Frmacc1155", strExec) = False Then
            Exit Sub
         End If
         tool1_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1155.Show
         Me.MousePointer = vbDefault
      Case 2 '收款作業
         If CheckUse("Frmacc1150", strExec) = False Then
            Exit Sub
         End If
         tool1_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1150.Show
         Me.MousePointer = vbDefault
         
      Case 3 '銷帳退費作業
         If CheckUse("Frmacc1190", strExec) = False Then
            Exit Sub
         End If
         tool1_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1190.Show
         Me.MousePointer = vbDefault
         
      Case 4 '暫收款作業
         If CheckUse("Frmacc11a0", strExec) = False Then
            Exit Sub
         End If
         tool1_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc11a0.Show
         Me.MousePointer = vbDefault
   End Select
End Sub

'Modified by Lydia 2023/11/13 調整國內應收之基本資料的選單
'*****國內應收->基本資料->發票*****
Private Sub Main1_1_1_3_Click(Index As Integer)
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   
   Select Case Index
      Case 1 '請款單開立發票作業
         If CheckUse("Frmacc1127", strExec) = False Then
            Exit Sub
         End If
         strItemNo = "" 'Add By Sindy 2017/2/14
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1127.Show
         Me.MousePointer = vbDefault
         
      Case 2 '發票作廢作業
         If CheckUse("Frmacc11q0", strExec) = False Then
            Exit Sub
         End If
         tool14_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc11q0.Show
         Me.MousePointer = vbDefault
         
      Case 3 '發票申報作業
         If CheckUse("Frmacc11s0", strExec) = False Then
            Exit Sub
         End If
         tool14_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc11s0.Show
         Me.MousePointer = vbDefault
         
      Case 4 '發票跨期轉開作業
         If CheckUse("Frmacc43b0", strExec) = False Then
            Exit Sub
         End If
         tool14_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc43b0.Show
         Me.MousePointer = vbDefault
         
      Case 5 '發票上傳作業
         If CheckUse("Frmacc11t0", strExec) = False Then
            Exit Sub
         End If
         tool14_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc11t0.Show
         Me.MousePointer = vbDefault
         
      Case 6 '發票號碼維護
        If CheckUse("Frmacc11o0", strExec) = False Then
            Exit Sub
         End If
         If Pub_StrUserSt03 = "M51" Then
            tool1_enabled '有刪除功能
         Else
            tool14_enabled
         End If
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc11o0.Show
         Me.MousePointer = vbDefault
         
      Case 7 '特殊發票客戶資料維護及查詢
         If CheckUse("Frmacc11o6", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc11o6.Show
         Me.MousePointer = vbDefault

      Case 8 '發票備註維護作業
        If CheckUse("Frmacc1128", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1128.Show
         Me.MousePointer = vbDefault
   End Select
End Sub

'Modified by Lydia 2023/11/13 調整國內應收之基本資料的選單
'*****國內應收->基本資料->其他*****
Private Sub Main1_1_1_4_Click(Index As Integer)
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   
   Select Case Index
      Case 1 '國內應收待處理作業
         If CheckUse("Frmacc11r0", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc11r0.Show
         Me.MousePointer = vbDefault
         
      Case 2 '收據抬頭基本資料維護
         If CheckUse("Frmacc11p0", strExec) = False Then
            Exit Sub
         End If
         tool1_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc11p0.Show
         Me.MousePointer = vbDefault
         
      Case 3 '退費收訖憑單維護
         If CheckUse("Frmacc11i0", strExec) = False Then
            Exit Sub
         End If
         tool1_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc11i0.Show
         Me.MousePointer = vbDefault
                  
      Case 4 '手開收據開立
         If CheckUse("Frmacc1110", strExec) = False Then
            Exit Sub
         End If
         tool9_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1110.Show
         Me.MousePointer = vbDefault
         
      Case 5 '收文金額分配作業
         If CheckUse("Frmacc11l0", strExec) = False Then
            Exit Sub
         End If
         tool1_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc11l0.Show
         Me.MousePointer = vbDefault
         
      Case 6 '客戶電匯資料維護及查詢
         If CheckUse("Frmacc11n1", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc11n1.Show
         Me.MousePointer = vbDefault
               
      Case 7    '客戶應收帳款收文檢查上限
            If CheckUse("frm140502", strExec) = True Then
               frm140502.Show
            End If
            
      Case 8    '客戶特殊付款週期維護
            If CheckUse("frm140504", strExec) = True Then
               frm140504.Show
            End If
   End Select
End Sub

'Mark by Lydia 2023/11/13 (舊表單) 調整國內應收之基本資料的選單
'Private Sub Main1_1_1_Click(Index As Integer)
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'
'   Select Case Index
'      Case 1 '收據開立作業
'         If CheckUse("Frmacc1120", strExec) = False Then
'            Exit Sub
'         End If
'         If PUB_GetLock("Frmacc1120", "", "收據開立作業") = False Then
'            Exit Sub
'         End If
'         tool3_enabled
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc1120.Show
'         Me.MousePointer = vbDefault
'
'      Case 2 '收據/請款單作廢作業
'         If CheckUse("Frmacc1130", strExec) = False Then
'            Exit Sub
'         End If
'         'Modify by Morgan 2006/3/20 取消刪除功能
'         'tool1_enabled
'         tool14_enabled
'         '2006/3/20 end
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc1130.Show
'         Me.MousePointer = vbDefault
'
'      Case 3 '收據抬頭修改
'         If CheckUse("Frmacc1140", strExec) = False Then
'            Exit Sub
'         End If
'         tool8_enabled 'Modify by Amy 2017/09/04 原:tool14_enabled
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc1140.Show
'         Me.MousePointer = vbDefault
'
'      Case 4 '收據金額修改
'         If CheckUse("Frmacc11d0", strExec) = False Then
'            Exit Sub
'         End If
'         'Modify by Morgan 2004/1/12
'         '加可按新增清除畫面
'         'tool8_enabled
'         tool14_enabled
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc11d0.Show
'         Me.MousePointer = vbDefault
'
'      Case 5 '收款作業
'         If CheckUse("Frmacc1150", strExec) = False Then
'            Exit Sub
'         End If
'         tool1_enabled
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc1150.Show
'         Me.MousePointer = vbDefault
'
'      Case 6 '銷帳退費作業
'         If CheckUse("Frmacc1190", strExec) = False Then
'            Exit Sub
'         End If
'         tool1_enabled
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc1190.Show
'         Me.MousePointer = vbDefault
'
'      Case 7 '暫收款作業
'         If CheckUse("Frmacc11a0", strExec) = False Then
'            Exit Sub
'         End If
'         tool1_enabled
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc11a0.Show
'         Me.MousePointer = vbDefault
'
'      Case 8 '手開收據開立
'         If CheckUse("Frmacc1110", strExec) = False Then
'            Exit Sub
'         End If
'         tool9_enabled
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc1110.Show
'         Me.MousePointer = vbDefault
'
'      Case 9 '退費收訖憑單維護
'         If CheckUse("Frmacc11i0", strExec) = False Then
'            Exit Sub
'         End If
'         tool1_enabled
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc11i0.Show
'         Me.MousePointer = vbDefault
'
'      'Add by Morgan 2010/12/6
'      Case 10 '收據開立作業-整批
'         If CheckUse("Frmacc1123", strExec) = False Then
'            Exit Sub
'         End If
'         If PUB_GetLock("Frmacc1123", "", "收據開立作業-批次") = False Then
'            Exit Sub
'         End If
'         tool3_enabled
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc1123.Show
'         Me.MousePointer = vbDefault
'
'      'Add by Morgan 2011/4/6
'      Case 11 '收文金額分配作業
'         If CheckUse("Frmacc11l0", strExec) = False Then
'            Exit Sub
'         End If
'         tool1_enabled
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc11l0.Show
'         Me.MousePointer = vbDefault
'
'      'Add by Morgan 2011/9/26
'      Case 12 '拆收據作業
'         If CheckUse("Frmacc11m0", strExec) = False Then
'            Exit Sub
'         End If
'         tool3_enabled
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc11m0.Show
'         Me.MousePointer = vbDefault
'
'      'Add by Sindy 2012/8/29
'      Case 13 '客戶電匯資料維護及查詢
'         If CheckUse("Frmacc11n1", strExec) = False Then
'            Exit Sub
'         End If
'         'tool1_enabled
'         tool3_enabled
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc11n1.Show
'         Me.MousePointer = vbDefault
'
'      'Add by Amy 2013/11/27
'      Case 14 '發票號碼維護
'        If CheckUse("Frmacc11o0", strExec) = False Then
'            Exit Sub
'         End If
'         If Pub_StrUserSt03 = "M51" Then
'            tool1_enabled '有刪除功能
'         Else
'            tool14_enabled
'         End If
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc11o0.Show
'         Me.MousePointer = vbDefault
'
'      'Add by Sindy 2013/12/13
'      Case 15 '特殊發票客戶資料維護及查詢
'         If CheckUse("Frmacc11o6", strExec) = False Then
'            Exit Sub
'         End If
'         'tool1_enabled
'         tool3_enabled
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc11o6.Show
'         Me.MousePointer = vbDefault
'      'Added by Morgan 2013/12/17
'      Case 16
'         If CheckUse("Frmacc1155", strExec) = False Then
'            Exit Sub
'         End If
'         tool1_enabled
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc1155.Show
'         Me.MousePointer = vbDefault
'      'Add by Sindy 2013/12/19
'      Case 17 '收據抬頭基本資料維護
'         If CheckUse("Frmacc11p0", strExec) = False Then
'            Exit Sub
'         End If
'         tool1_enabled
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc11p0.Show
'         Me.MousePointer = vbDefault
'      'Add by Sindy 2013/12/31
'      Case 18 '請款單開立發票作業
'         If CheckUse("Frmacc1127", strExec) = False Then
'            Exit Sub
'         End If
'         strItemNo = "" 'Add By Sindy 2017/2/14
'         tool3_enabled
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc1127.Show
'         Me.MousePointer = vbDefault
'      'Add by Amy 2022/08/24
'      Case 19 '發票備註維護作業
'        If CheckUse("Frmacc1128", strExec) = False Then
'            Exit Sub
'         End If
'         tool3_enabled
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc1128.Show
'         Me.MousePointer = vbDefault
'      'Add by Sindy 2014/1/2
'      Case 20 '國內應收待處理作業
'         If CheckUse("Frmacc11r0", strExec) = False Then
'            Exit Sub
'         End If
'         tool3_enabled
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc11r0.Show
'         Me.MousePointer = vbDefault
'      'Add by Sindy 2014/1/9
'      Case 21 '發票作廢作業
'         If CheckUse("Frmacc11q0", strExec) = False Then
'            Exit Sub
'         End If
'         tool14_enabled
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc11q0.Show
'         Me.MousePointer = vbDefault
'      'Add by Sonia 2014/3/6
'      Case 22 '發票申報作業
'         If CheckUse("Frmacc11s0", strExec) = False Then
'            Exit Sub
'         End If
'         tool14_enabled
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc11s0.Show
'         Me.MousePointer = vbDefault
'      'Add by Sonia 2014/3/12
'      Case 23 '發票跨期轉開作業
'         If CheckUse("Frmacc43b0", strExec) = False Then
'            Exit Sub
'         End If
'         tool14_enabled
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc43b0.Show
'         Me.MousePointer = vbDefault
'     'Add byAmy 2019/03/05
'      Case 24 '發票上傳作業
'         If CheckUse("Frmacc11t0", strExec) = False Then
'            Exit Sub
'         End If
'         tool14_enabled
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc11t0.Show
'         Me.MousePointer = vbDefault
'      'Added by Lydia 2020/02/07
'      Case 25    '客戶應收帳款收文檢查上限
'            If CheckUse("frm140502", strExec) = True Then
'               frm140502.Show
'            End If
'      Case 26    '客戶特殊付款週期維護
'            If CheckUse("frm140504", strExec) = True Then
'               frm140504.Show
'            End If
'   End Select
'End Sub
'end 2023/11/13

'收據資料查詢
Private Sub Main1_2_1_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc1211", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc1211.Show
   Me.MousePointer = vbDefault
End Sub

'Add By Sindy 2010/4/19 未列印收據/請款單查詢
Private Sub Main1_2_10_Click()
   Toolbar1.Visible = False
   StatusBar1.Visible = False
   MenuDisabled
   frm210132.Show
End Sub

'發票作廢查詢 ADD BY ERIC 2014/1/22
Private Sub Main1_2_11_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc12c0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc12c0.Show
   Me.MousePointer = vbDefault
End Sub

'Add By Sindy 2014/1/23 已開發票未收款明細查詢
Private Sub Main1_2_12_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc14n0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc14n0.Show
   Me.MousePointer = vbDefault
End Sub

'Added by Sindy 2014/3/11
Private Sub Main1_2_13_Click()
   Toolbar1.Visible = False
   StatusBar1.Visible = False
   MenuDisabled
   frm210143.Show
End Sub

'Add By Sindy 2016/6/6
'應收帳款綜合查詢
Private Sub Main1_2_14_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc12d0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc12d0.Show
   Me.MousePointer = vbDefault
End Sub

'Add By Sindy 2020/9/21 法律與智慧所案件對照表
Private Sub Main1_2_15_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc12e0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc12e0.Show
   Me.MousePointer = vbDefault
End Sub

'Added by Morgan 2024/8/12
'收文金額異常檢查
Private Sub Main1_2_16_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc12f0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc12f0.Show
   Me.MousePointer = vbDefault
End Sub


'客戶帳款查詢
Private Sub Main1_2_2_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc1220", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc1220.Show
   Me.MousePointer = vbDefault
End Sub

'智權人員帳款查詢
Private Sub Main1_2_3_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc1230", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc1230.Show
   Me.MousePointer = vbDefault
End Sub

'本所案號帳目查詢
Private Sub Main1_2_4_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc1240", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc1240.Show
   Me.MousePointer = vbDefault
End Sub

'收款單號查詢
Private Sub Main1_2_5_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc1250", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc1250.Show
   Me.MousePointer = vbDefault
End Sub

'收據/請款單作廢查詢
Private Sub Main1_2_6_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc1260", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc1260.Show
   Me.MousePointer = vbDefault
End Sub

'收文與收據資料檢核查詢
Private Sub Main1_2_7_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc1280", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc1280.Show
   Me.MousePointer = vbDefault
End Sub

'手開收據資料查詢
Private Sub Main1_2_8_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc1290", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc1290.Show
   Me.MousePointer = vbDefault
End Sub

'Add by Morgan 2010/2/3
'客戶應收帳款查詢
Private Sub Main1_2_9_Click()
   Toolbar1.Visible = False
   StatusBar1.Visible = False
   MenuDisabled
   frm210122.cmdEdit.Enabled = False
   frm210122.Show
End Sub


Private Sub Main1_4_1_Click(Index As Integer)
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   
   Select Case Index
      Case 1 '收據列印
         If CheckUse("Frmacc1410", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1410.ProState = "1" '權限: 1.全所 2.該所 Add By Sindy 2021/5/21
         Frmacc1410.Show
         Me.MousePointer = vbDefault
      Case 2 '補開收據列印
         If CheckUse("Frmacc1420", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1420.ProState = "1" '權限: 1.全所 2.該所 Add By Sindy 2021/5/21
         Frmacc1420.Show
         Me.MousePointer = vbDefault
      Case 3 '客戶對帳單
         If CheckUse("Frmacc1440", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1440.Show
         Me.MousePointer = vbDefault
      Case 4 '客戶帳款明細表
         If CheckUse("Frmacc1450", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1450.Show
         Me.MousePointer = vbDefault
      Case 5 '智權人員帳款明細表
         If CheckUse("Frmacc1460", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1460.Show
         Me.MousePointer = vbDefault
      Case 6 '智權人員帳款明細表員應收規費明細表
         tool3_enabled
         If CheckUse("Frmacc1470", strExec) = False Then
            Exit Sub
         End If
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1470.Show
         Me.MousePointer = vbDefault
      Case 7 '國內帳齡分析表
         If CheckUse("Frmacc1480", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1480.Show
         Me.MousePointer = vbDefault
      'Modify by Amy 2017/03/16 加 國內人員收文點數及收款點數統計
      Case 8
         If CheckUse("Frmacc14w0", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc14w0.Show
         Me.MousePointer = vbDefault
      Case 9 '銷帳退費明細表
         If CheckUse("Frmacc1490", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1490.Show
         Me.MousePointer = vbDefault
      Case 10 '暫收款明細表
         If CheckUse("Frmacc14a0", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc14a0.Show
         Me.MousePointer = vbDefault
      Case 11 '收據/請款單作廢明細表
         If CheckUse("Frmacc14e0", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc14e0.Show
         Me.MousePointer = vbDefault
      Case 12 '收文與收據資料檢核表
         If CheckUse("Frmacc14f0", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc14f0.Show
         Me.MousePointer = vbDefault
      Case 13 '收據抬頭修改清單
         If CheckUse("Frmacc14h0", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc14h0.Show
         Me.MousePointer = vbDefault
      'Add by Morgan 2011/10/3
      Case 14 '收據帳款明細列印
         If CheckUse("Frmacc14l0", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc14l0.Show
         Me.MousePointer = vbDefault
      'Add By Sindy 2012/4/27
      Case 15 '收款扣繳改年度清單
         If CheckUse("Frmacc14m0", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc14m0.Show
         Me.MousePointer = vbDefault
      'Add By Sindy 2013/12/3
      Case 16 '請款單及發票列印     '2024/4/22 sonia改為請款單列印
         If CheckUse("Frmacc1610", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1610.ProState = "1" '權限: 1.全所 2.該所 Add By Sindy 2021/5/21
         Frmacc1610.Show
         Me.MousePointer = vbDefault
      Case 17 '補開請款單及發票列印 '2024/4/22 sonia改為補開請款單列印
         If CheckUse("Frmacc1620", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1620.ProState = "1" '權限: 1.全所 2.該所 Add By Sindy 2021/5/21
         Frmacc1620.Show
         Me.MousePointer = vbDefault
      Case 18 '同仁介紹案源獎金明細表
         If CheckUse("Frmacc14t0", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc14t0.Show
         Me.MousePointer = vbDefault
      'Modify By Sindy 2022/3/24 辜苑琪(Sent: Thursday, March 24, 2022 10:57 AM)說: 已改至盟立系統, 這支完全不用了
'      'Add By Sindy 2014/4/3
'      Case 19 '銷貨退回折讓單列印
'         If CheckUse("Frmacc1630", strExec) = False Then
'            Exit Sub
'         End If
'         tool3_enabled
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc1630.Show
'         Me.MousePointer = vbDefault
      'Add By Sindy 2014/4/22
      Case 20 '國內收據產生特殊請款單
         If CheckUse("Frmacc14p0", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc14p0.Show
         Me.MousePointer = vbDefault
      'Add By Sindy 2014/5/7
      'Mark by Lydia 2023/11/13 改到基本資料->收據;後面Index-1
'      Case 21 '國內收據產生INVOICE
'         If CheckUse("Frmacc14o0", strExec) = False Then
'            Exit Sub
'         End If
'         tool3_enabled
'         Me.MousePointer = vbHourglass
'         MenuDisabled
'         Frmacc14o0.Show
'         Me.MousePointer = vbDefault
      'end 2023/11/13
      'Add By Amy 2015/04/07
      Case 21 '應收帳款財務處控管資料表
      'end 2017/03/16
         If CheckUse("Frmacc14u0", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc14u0.Show
         Me.MousePointer = vbDefault
      'Added by Lydia 2019/10/04
      Case 22 '智權人員請款明細表
         frm210146.Show
   End Select
End Sub

Private Sub Main1_Click()
   Toolbar1.Visible = True
   StatusBar1.Visible = True
End Sub

'收款作業
Private Sub Main2_1_1_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc2110", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc2110.Show
   Me.MousePointer = vbDefault
End Sub

'結匯匯率輸入
Private Sub Main2_1_10_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc21c0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc21c0.Show
   Me.MousePointer = vbDefault
End Sub

'匯票輸入
Private Sub Main2_1_11_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc21d0", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc21d0.Show
   Me.MousePointer = vbDefault
End Sub

'付款後退費作業
Private Sub Main2_1_12_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc21e0", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc21e0.Show
   Me.MousePointer = vbDefault
End Sub

'抵帳作業
Private Sub Main2_1_13_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc21f0", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc21f0.Show
   Me.MousePointer = vbDefault
End Sub

'外幣票據Excel檔案產生
'Memo by Lydia 2018/07/12 更名為"國外收款分析表"
Private Sub Main2_1_18_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc21l0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc21l0.Show
   Me.MousePointer = vbDefault
End Sub

'銀存匯率輸入
Private Sub Main2_1_19_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc21n0", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc21n0.Show
   Me.MousePointer = vbDefault
End Sub

'暫收款作業
Private Sub Main2_1_2_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc2120", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc2120.Show
   Me.MousePointer = vbDefault
End Sub

'客戶/代理人匯款銀行資料維護
Private Sub Main2_1_20_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc21q0", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc21q0.Show
   Me.MousePointer = vbDefault
End Sub

'客戶/代理人財務EMail資料維護
Private Sub Main2_1_21_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc21r0", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc21r0.Show
   Me.MousePointer = vbDefault
End Sub

'暫收款退費作業
Private Sub Main2_1_3_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc2130", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc2130.Show
   Me.MousePointer = vbDefault
End Sub

'銷帳作業
Private Sub Main2_1_4_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc2140", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc2140.Show
   Me.MousePointer = vbDefault
End Sub

'結匯資料輸入
Private Sub Main2_1_7_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc2170", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc2170.Show
   Me.MousePointer = vbDefault
End Sub

'調整付款明細
Private Sub Main2_1_9_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc21b0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc21b0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main2_2_1_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc2210", strExec) = False Then
      Exit Sub
   End If
   'Modify by Amy 2015/08/26 開放放大鏡鈕
   'tool3_enabled
   tool13_enabled
   'end 2015/08/26
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc2210.Show
   Me.MousePointer = vbDefault
End Sub

'Add by Amy 2021/05/04
Private Sub Main2_2_10_Click()
    If strFormName <> MsgText(601) Then
        Exit Sub
    End If
    tool3_enabled
    Me.MousePointer = vbHourglass
    MenuDisabled
    Frmacc24n0.Show
    Me.MousePointer = vbDefault
End Sub

Private Sub Main2_2_2_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc2220", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc2220.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main2_2_3_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc2230", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc2230.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main2_2_4_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc2250", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc2250.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main2_2_5_Click()
   Toolbar1.Visible = False
   StatusBar1.Visible = False
    'Add By Cheng 2004/05/18
    If CheckUse("frm050203", strExec) = False Then
        Exit Sub
    End If
    frm050203.Show
    'End
End Sub
'2008/12/4 add by sonia
Private Sub Main2_2_6_Click()
   Toolbar1.Visible = False
   StatusBar1.Visible = False
   StrStartSystemByNick = "ALL"
    If CheckUse("frm040205", strExec) = False Then
        Exit Sub
    End If
    frm040205.Show
End Sub
'2008/12/4 end

'Add By Sindy 2010/3/26
Private Sub Main2_2_7_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
'   If CheckUse("Frmacc2260", strExec) = False Then
'      Exit Sub
'   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc2260.Show
   Me.MousePointer = vbDefault
End Sub

'Add by Amy 20130913 各幣別最新請款匯率查詢
Private Sub Main2_2_8_Click()
    If CheckUse("Frmacc2142", strExec) = False Then
        Exit Sub
    End If
    Toolbar1.Visible = False
    StatusBar1.Visible = False
    Frmacc2142.Show
End Sub

'Add by Amy 2021/03/04 FC業務請款／收款明細表(國外部用)
Private Sub Main2_2_9_Click()
    If strFormName <> MsgText(601) Then
        Exit Sub
    End If
    tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc24c0.Show
   Me.MousePointer = vbDefault
End Sub

'請款明細刪除作業 2019/10/29 Sonia設定不顯示
Private Sub Main2_3_1_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc2310", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc2310.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main2_4_10_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc24a0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc24a0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main2_4_13_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc24d0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc24d0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main2_4_14_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc24e0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc24e0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main2_4_15_Click()
   If CheckUse("frm040329", strExec) = False Then
      Exit Sub
   End If
   frm040329.Show
End Sub

Private Sub Main2_4_16_Click()
   If CheckUse("frm083014", strExec) = False Then
      Exit Sub
   End If
   frm083014.Show
End Sub

'Remove by Lydia 2017/02/24 未使用
'Private Sub Main2_4_17_Click()
'   If CheckUse("frmacc24j0", strExec) = False Then
'      Exit Sub
'   End If
'   Frmacc24j0.Show
'End Sub

'Add By Sindy 2009/10/02
Private Sub Main2_4_18_Click()
   '國外部FC案件不請款清單
   If CheckUse("frm040332", strExec) = False Then
      Exit Sub
   End If
   frm040332.Show
End Sub

Private Sub Main2_4_19_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc24k0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc24k0.Show
   Me.MousePointer = vbDefault
End Sub
'Add by Lydia 2015/02/16 台銀結匯水單媒體產生作業
'Memo by Lydia 2017/09/26 更名為:國外結匯媒體檔產生作業
Private Sub Main2_4_20_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc24m0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc24m0.Show
   Me.MousePointer = vbDefault
End Sub

'Add By Sindy 2015/10/20
Private Sub Main2_4_21_Click()
   If CheckUse("Frmacc14v0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc14v0.Show
   Me.MousePointer = vbDefault
End Sub

'Memo by Lydia 2015/04/17 原名"水單列印",更名為"非台銀媒體水單列印"
'Memo by Lydia 2017/09/26 更名為:國外結匯水單列印
Private Sub Main2_4_3_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc2430", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc2430.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main2_4_4_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc2440", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc2440.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main2_4_5_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc2450", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc2450.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main2_4_6_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc2460", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc2460.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main2_4_7_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc2470", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc2470.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main2_4_8_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc2480", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc2480.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main2_4_9_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc2490", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc2490.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main2_Click()
   Toolbar1.Visible = True
   StatusBar1.Visible = True
End Sub

Private Sub Main3_1_1_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc3110", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc3110.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_1_10_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc31a0", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc31a0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_1_11_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc31b0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc31b0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_1_12_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc31c0", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc31c0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_1_13_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc31d0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc31d0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_1_14_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc31e0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc31e0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_1_15_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc31f0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc31f0.Show
   Me.MousePointer = vbDefault
End Sub
'Add by Morgan 2011/6/2
Private Sub Main3_1_16_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc31g0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc31g0.Show
   Me.MousePointer = vbDefault
End Sub
'Add by Morgan 2012/10/11
Private Sub Main3_1_17_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc31h0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc31h0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_1_2_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc3120", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc3120.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_1_3_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc3130", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc3130.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_1_4_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc3140", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc3140.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_1_5_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc3150", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc3150.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_1_6_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc3160", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc3160.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_1_7_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc3170", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc3170.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_1_8_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc3180", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc3180.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_1_9_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc3190", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc3190.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_2_1_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc3210", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc3210.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_2_10_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc32a0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc32a0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_2_11_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc32b0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc32b0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_2_12_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc32c0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc32c0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_2_13_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc32g0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc32g0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_2_14_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc32h0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc32h0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_2_15_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc32f0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc32f0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_2_2_Click()
   'Mark by Amy 2022/02/23 過濾Form2.0 未使用表單,發現此功能未顯示,故先刪
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc3220", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc3220.Show     '在功能表未顯示
'   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_2_3_Click()
   'Mark by Amy 2022/02/23 過濾Form2.0 未使用表單,發現此功能未顯示,故先刪
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc3230", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc3230.Show     '在功能表未顯示
'   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_2_5_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc3250", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc3250.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_2_6_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc3260", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc3260.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_2_7_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc3270", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc3270.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_2_8_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc3280", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc3280.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_3_1_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc3310", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc3310.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_3_2_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc3320", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc3320.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_4_1_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc3410", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc3410.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_4_10_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc34a0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc34a0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_4_11_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc34b0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc34b0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_4_12_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc34c0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc34c0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_4_13_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc34d0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc34d0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_4_14_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc34e0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc34e0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_4_16_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc34g0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc34g0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_4_17_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc34h0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc34h0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_4_18_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc34i0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc34i0.Show     '在功能表未顯示
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_4_19_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc34j0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc34j0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_4_2_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc3420", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc3420.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_4_3_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc3430", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc3430.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_4_4_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc3440", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc3440.Show     '在功能表未顯示
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_4_5_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc3450", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc3450.Show     '在功能表未顯示
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_4_6_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc3460", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc3460.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_4_7_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc3470", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc3470.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_4_8_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc3480", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc3480.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_4_9_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc3490", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc3490.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_Click()
   Toolbar1.Visible = True
   StatusBar1.Visible = True
End Sub

'Remove by Lydia 2017/02/24 改成Main4_1_1
'Private Sub Main4_1_10_Click()
'   'edit by nickc 2007/02/08
'   'If strform <> MsgText(601) Then
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc41a0", strExec) = False Then
'      Exit Sub
'   End If
'   tool1_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc41a0.Show
'   Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_1_1
'Private Sub Main4_1_11_Click()
'   'edit by nickc 2007/02/08
'   'If strform <> MsgText(601) Then
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc41b0", strExec) = False Then
'      Exit Sub
'   End If
'   tool1_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc41b0.Show
'   Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_1_1
'Private Sub Main4_1_12_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc41c0", strExec) = False Then
'      Exit Sub
'   End If
'   'Modify by Morgan 2004/10/27
'   'tool3_enabled
'   'Modify  by Amy 2014/02/20 +先輸入公司別
'   tool4_enabled
'   Frmacc41c2.Show vbModal
'   If strCompanyNo <> "" Then
'        tool6_enabled
'        Me.MousePointer = vbHourglass
'        MenuDisabled
'        Frmacc41c0.Show
'        Me.MousePointer = vbDefault
'   End If
'   'end 2014/02/20
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_1_1
'Private Sub Main4_1_13_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc41d0", strExec) = False Then
'      Exit Sub
'   End If
'   'Modify by Amy 2014/02/11
'   'tool1_enabled
'   tool8_enabled
'   'end 2014/02/11
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc41d0.Show
'   Me.MousePointer = vbDefault
'End Sub

'Add by Morgan 2005/4/6
'Remove by Lydia 2017/02/24 改成Main4_1_1
'Private Sub Main4_1_14_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc41e0", strExec) = False Then
'      Exit Sub
'   End If
'   tool1_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc41e0.Show
'   Me.MousePointer = vbDefault
'End Sub

'add by nickc 2005/11/10 加入結餘維護
'Remove by Lydia 2017/02/24 改成Main4_1_1
'Private Sub Main4_1_15_Click()
'   If CheckUse("frm040206", strExec) = True Then
'      frm040206.Show
'   End If
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_1_1
'Private Sub Main4_1_16_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   'If CheckUse("frmacc41f0", strExec) = false Then
'      'Exit Sub
'   'End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc41f0.Show
'   Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_1_1
'Private Sub Main4_1_17_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4170_1", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4170_1.Show
'   Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_1_1
'Private Sub Main4_1_18_Click()
'   Toolbar1.Visible = False
'   StatusBar1.Visible = False
'   MenuDisabled
'   frm210152.Show
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_1_1
'Private Sub Main4_1_19_Click()
'    'Add by Amy 2016/03/03
'    Me.MousePointer = vbHourglass
'    tool3_enabled
'    MenuDisabled
'    Frmacc43c0.Show
'    Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_1_1
'Private Sub Main4_1_2_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4120", strExec) = False Then
'      Exit Sub
'   End If
'   'Modify by Amy 2014/02/20 改從frmacc41c2輸入公司別
'   'Add by Amy 2014/01/07 +登入前輸入公司別
''   Call Frmacc4120.ChgCompany
''   If strExc(0) = "" Then
''        Exit Sub
''   End If
'   'end 2014/01/07
'   tool4_enabled
'   Frmacc41c2.Show vbModal
'   If strCompanyNo <> "" Then
'        tool1_enabled
'        Me.MousePointer = vbHourglass
'        MenuDisabled
'        'strCompanyNo = UCase(strExc(0)) 'Add by Amy 2014/01/07
'        Frmacc4120.Show
'        Me.MousePointer = vbDefault
'   End If
'   'end 2014/02/20
'End Sub

'Add by Amy 2016/07/11
'Remove by Lydia 2017/02/24
'Private Sub Main4_1_20_Click()
'    Me.MousePointer = vbHourglass
'    tool3_enabled
'    MenuDisabled
'    'Frmacc41j0.Show
'    Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_1_1
'Private Sub Main4_1_21_Click()
'    Me.MousePointer = vbHourglass
'    tool3_enabled
'    MenuDisabled
'    Frmacc41g0.Show
'    Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_1_1
'Private Sub Main4_1_22_Click()
'    Me.MousePointer = vbHourglass
'    tool3_enabled
'    MenuDisabled
'    Frmacc41h0.Show
'    Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_1_1
'Private Sub Main4_1_3_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4130", strExec) = False Then
'      Exit Sub
'   End If
'   tool1_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4130.Show
'   Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_1_1
'Private Sub Main4_1_4_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4140", strExec) = False Then
'      Exit Sub
'   End If
'   tool1_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4140.Show
'   Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_1_1
'Private Sub Main4_1_6_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4160", strExec) = False Then
'      Exit Sub
'   End If
'   tool1_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4160.Show
'   Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_1_1
'Private Sub Main4_1_7_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4170", strExec) = False Then
'      Exit Sub
'   End If
'   tool1_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4170.Show
'   Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_1_1
'Private Sub Main4_1_8_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4180", strExec) = False Then
'      Exit Sub
'   End If
'   tool1_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4180.Show
'   Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_1_1
'Private Sub Main4_1_9_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4190", strExec) = False Then
'      Exit Sub
'   End If
'   tool1_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4190.Show
'   Me.MousePointer = vbDefault
'End Sub

'Modified by Lydia 2017/02/24 改成Index
'Private Sub Main4_1_1_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4110", strExec) = False Then
'      Exit Sub
'   End If
'   tool1_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4110.Show
'   Me.MousePointer = vbDefault
'End Sub
Private Sub Main4_1_1_Click(Index As Integer)
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   
   Select Case Index
      Case 1 '會計科目基本資料
        If CheckUse("Frmacc4110", strExec) = False Then
           Exit Sub
        End If
        tool1_enabled
        Me.MousePointer = vbHourglass
        MenuDisabled
        Frmacc4110.Show
        Me.MousePointer = vbDefault
        
      Case 2 '傳票輸入
        If CheckUse("Frmacc4120", strExec) = False Then
           Exit Sub
        End If
        'Add by Amy 2017/10/25 判斷與自動轉傳票程式不可同時執行
        If TranNoLock("Frmacc4120") = False Then
            Exit Sub
        End If
        tool4_enabled
        Frmacc41c2.Show vbModal
        If strCompanyNo <> "" Then
            'Modify by Amy 2024/07/31 程式已判斷傳票不連號問題,故開放都可使用刪除
'            'Modify by Amy 2024/02/05 M51可使用刪除
'            If Pub_StrUserSt03 = "M51" Then
               tool1_enabled
'            Else
'               'Modify by Amy 2023/12/06 避免傳票不連號,取消刪除鈕  原:tool1_enabled
'               tool14_enabled
'            End If
             Me.MousePointer = vbHourglass
             MenuDisabled
             Frmacc4120.Show
             Me.MousePointer = vbDefault
        End If
        
      Case 3 '公司基本資料
        If CheckUse("Frmacc4130", strExec) = False Then
           Exit Sub
        End If
        tool1_enabled
        Me.MousePointer = vbHourglass
        MenuDisabled
        Frmacc4130.Show
        Me.MousePointer = vbDefault
        
      Case 4 '部門基本資料
        If CheckUse("Frmacc4140", strExec) = False Then
           Exit Sub
        End If
        tool1_enabled
        Me.MousePointer = vbHourglass
        MenuDisabled
        Frmacc4140.Show
        Me.MousePointer = vbDefault
        
      Case 5 '預算資料
        If CheckUse("Frmacc4160", strExec) = False Then
           Exit Sub
        End If
        tool1_enabled
        Me.MousePointer = vbHourglass
        MenuDisabled
        Frmacc4160.Show
        Me.MousePointer = vbDefault
        
      Case 6 '每月固定傳票資料
        If CheckUse("Frmacc4170", strExec) = False Then
           Exit Sub
        End If
        tool1_enabled
        Me.MousePointer = vbHourglass
        MenuDisabled
        Frmacc4170.Show
        Me.MousePointer = vbDefault
        
      Case 7 '分攤類別資料
        If CheckUse("Frmacc4180", strExec) = False Then
           Exit Sub
        End If
        tool1_enabled
        Me.MousePointer = vbHourglass
        MenuDisabled
        Frmacc4180.Show
        Me.MousePointer = vbDefault
        
      Case 8 '分攤類別比率資料
        If CheckUse("Frmacc4190", strExec) = False Then
           Exit Sub
        End If
        tool1_enabled
        Me.MousePointer = vbHourglass
        MenuDisabled
        Frmacc4190.Show
        Me.MousePointer = vbDefault
        
      Case 9 'CF案件結餘結算作業
        If CheckUse("Frmacc41a0", strExec) = False Then
           Exit Sub
        End If
        tool1_enabled
        Me.MousePointer = vbHourglass
        MenuDisabled
        Frmacc41a0.Show
        Me.MousePointer = vbDefault
        
      Case 10 'CF案件結餘作廢作業
        If CheckUse("Frmacc41b0", strExec) = False Then
           Exit Sub
        End If
        tool1_enabled
        Me.MousePointer = vbHourglass
        MenuDisabled
        Frmacc41b0.Show
        Me.MousePointer = vbDefault
'cancel by sonia 2023/7/20 移至查詢作業
'      Case 11 'CF案件結餘維護
'        If CheckUse("frm040206", strExec) = True Then
'           frm040206.Show
'        End If
'end 2023/7/20
        
      Case 12 '傳票過帳後摘要修改
        If CheckUse("Frmacc41c0", strExec) = False Then
           Exit Sub
        End If
        tool4_enabled
        Frmacc41c2.Show vbModal
        If strCompanyNo <> "" Then
             tool6_enabled
             Me.MousePointer = vbHourglass
             MenuDisabled
             Frmacc41c0.Show
             Me.MousePointer = vbDefault
        End If
        
      Case 13 '應收付分錄調整
        If CheckUse("Frmacc41d0", strExec) = False Then
           Exit Sub
        End If
        tool8_enabled
        Me.MousePointer = vbHourglass
        MenuDisabled
        Frmacc41d0.Show
        Me.MousePointer = vbDefault
        
      Case 14 '簽收作業
        If CheckUse("Frmacc41e0", strExec) = False Then
           Exit Sub
        End If
        tool1_enabled
        Me.MousePointer = vbHourglass
        MenuDisabled
        Frmacc41e0.Show
        Me.MousePointer = vbDefault
        
      Case 15 '結餘保留放出產生傳票
        tool3_enabled
        Me.MousePointer = vbHourglass
        MenuDisabled
        Frmacc41f0.Show
        Me.MousePointer = vbDefault
        
      Case 16 '每月業績開放/關閉輸入
        Me.MousePointer = vbHourglass
        tool3_enabled
        MenuDisabled
        Frmacc43c0.Show
        Me.MousePointer = vbDefault
        
      Case 17 '智權點數實績與結餘輸入
        Toolbar1.Visible = False
        StatusBar1.Visible = False
        MenuDisabled
        frm210152.IsAgentLimit = False 'Add by Amy 2023/02/03 +職代
        frm210152.Show
      'Modify by Amy 2017/10/02 修改順序
      'Modify by Amy 2017/06/15
      
      'Modify by Amy 2023/04/18 +ACS 待分潤
      Case 18 'ACS 待分潤
        If TranNoLock("Frmacc41l0") = False Then
            Exit Sub
        End If
        Me.MousePointer = vbHourglass
        tool3_enabled
        MenuDisabled
        Frmacc41l0.Show
        Me.MousePointer = vbDefault
        
      'Modify by Amy 2024/07/15 非智權結餘轉撥報出傳票產生(隱藏版)-改名稱 婉莘
      Case 19 '非當月結餘轉撥傳票產生(隱藏版人員)
        'Add by Amy 2017/10/25 判斷與自動轉傳票程式不可同時執行
        If TranNoLock("Frmacc41j0") = False Then
            Exit Sub
        End If
        Me.MousePointer = vbHourglass
        tool1_enabled
        MenuDisabled
        Frmacc41j0.Show
        Me.MousePointer = vbDefault
        
      Case 20 '智權期末實績保留傳票產生
        'Add by Amy 2017/10/25 判斷與自動轉傳票程式不可同時執行
        If TranNoLock("Frmacc41g0") = False Then
            Exit Sub
        End If
        Me.MousePointer = vbHourglass
        tool3_enabled
        MenuDisabled
        Frmacc41g0.Show
        Me.MousePointer = vbDefault
      Case 21 '智權期末結餘保留傳票產生
        'Add by Amy 2017/10/25 判斷與自動轉傳票程式不可同時執行
        If TranNoLock("Frmacc41h0") = False Then
            Exit Sub
        End If
        Me.MousePointer = vbHourglass
        tool3_enabled
        MenuDisabled
        Frmacc41h0.Show
        Me.MousePointer = vbDefault
      'end 2017/06/15
      'end 2017/10/02
      'Modify by Amy 2021/09/17 智權期末結餘保留資料刪除
      Case 22
        If TranNoLock("Frmacc41k0", "Frmacc41g0") = False Then
           Exit Sub
        End If
        tool3_enabled
        Me.MousePointer = vbHourglass
        MenuDisabled
        Frmacc41k0.Show
        Me.MousePointer = vbDefault
      Case 23 '每月固定傳票資料-非分攤
      'end 2023/02/18
      'end 2021/09/17
        If CheckUse("Frmacc4170_1", strExec) = False Then
           Exit Sub
        End If
        'Add by Amy 2022/06/07 判斷與自動轉傳票程式不可同時執行
         If TranNoLock("Frmacc4170_1") = False Then
            Exit Sub
         End If
        tool3_enabled
        Me.MousePointer = vbHourglass
        MenuDisabled
        Frmacc4170_1.Show
        Me.MousePointer = vbDefault
   End Select
End Sub
'end 2017/02/24

'Modified by Lydia 2017/02/24 改成Index
'Private Sub Main4_2_1_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4210", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4210.Show
'   Me.MousePointer = vbDefault
'End Sub
Private Sub Main4_2_1_Click(Index As Integer)
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   Select Case Index
      Case 1 '傳票資料查詢
        If CheckUse("Frmacc4210", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc4210.Show

      Case 2 '日記帳查詢
        If CheckUse("Frmacc4280", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc4280.Show

      Case 3 '科目分類帳查詢
        If CheckUse("Frmacc4220", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc4220.Show
      
      Case 4 '科目餘額查詢
        If CheckUse("Frmacc4230", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc4230.Show
        
      Case 5 '科目明細查詢(對沖)
        If CheckUse("Frmacc4240", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc4240.Show

      Case 6 '智權人員點數查詢
        If CheckUse("Frmacc4250", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc4250.Show

      Case 7 '單據/傳票號碼查詢
        If CheckUse("Frmacc4260", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc4260.Show

      Case 8 '智權人員結餘點數查詢
        If CheckUse("Frmacc4270", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc4270.Show

      Case 9 '簽收資料查詢
        Toolbar1.Visible = False
        StatusBar1.Visible = False
        MenuDisabled
        If frm210106_1.setNextForm = "" Then
           frm210106.Show
        Else
           frm210106_1.setCaller frm210106
           frm210106_1.Show
        End If
        
      Case 10 '業績點數查詢
        Toolbar1.Visible = False
        StatusBar1.Visible = False
        MenuDisabled
        frm210104.Show
        
      Case 11 'CF結餘單查詢
        If CheckUse("frm040202", strExec) = True Then
           frm040202.Show
        End If
        
      'add by 2019/11/29(共同程序2008/04上線)
      Case 12   'CF 結餘單案件明細查詢
         If CheckUse("frm040208", strExec) = True Then
            frm040208.Show
         End If
      'end 2019/11/29
      
'add by sonia 2023/7/26 從基本資料移過來並改財務處看到的Caption
      Case 13 'CF 可結餘日期查詢(專業部-CF 結餘資料維護)
        If CheckUse("frm040206", strExec) = True Then
           frm040206.Show
        End If
'end 2023/7/26
      
      Case 14 '專業點數分析查詢與列印
        If CheckUse("Frmacc42a0", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc42a0.Show

      Case 15 '未繳款簽收資料查詢
        If CheckUse("Frmacc42b0", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc42b0.Show

      Case 16 '過帳前綜合損益查詢及列印
        If CheckUse("Frmacc4290", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc4290.Show
        
      'Add By Sindy 2020/3/10
      Case 17 '旅遊補助付款通知
        If CheckUse("frmacc42c0", strExec) = False Then
           Exit Sub
        End If
        frmacc42c0.Show
      'Added by Lydia 2024/07/29
      Case 18 '出庭費發放通知 (113/11/01上線)
        If CheckUse("frmacc42d0", strExec) = False Then
           Exit Sub
        End If
        Frmacc42d0.Show

   End Select
   Me.MousePointer = vbDefault
End Sub
'end 2017/02/24

''Add by Morgan 2005/6/17 業績點數查詢
''Remove by Lydia 2017/02/24 改成Main4_2_1
'Private Sub Main4_2_10_Click()
'   Toolbar1.Visible = False
'   StatusBar1.Visible = False
'   MenuDisabled
'   frm210104.Show
'End Sub
'
''案件結餘查詢
''2006/2/14 ADD BY SONIA 加CF結餘查詢
''Remove by Lydia 2017/02/24 改成Main4_2_1
'Private Sub Main4_2_11_Click()
'   If CheckUse("frm040202", strExec) = True Then
'      frm040202.Show
'   End If
'End Sub
'
''專業點數分析查詢與列印
''Remove by Lydia 2017/02/24 改成Main4_2_1
'Private Sub Main4_2_12_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc42a0", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc42a0.Show
'   Me.MousePointer = vbDefault
'End Sub

'Add by Amy 2014/04/16 未繳款簽收資料查詢
'Remove by Lydia 2017/02/24 改成Main4_2_1
'Private Sub Main4_2_13_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc42b0", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc42b0.Show
'   Me.MousePointer = vbDefault
'End Sub

'Added by Lydia 2014/12/11 總帳-過帳前綜合損益查詢及列印
'Remove by Lydia 2017/02/24 改成Main4_2_1
'Private Sub Main4_2_14_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4290", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4290.Show
'   Me.MousePointer = vbDefault
'End Sub

'科目分類帳查詢
'Remove by Lydia 2017/02/24 改成Main4_2_1
'Private Sub Main4_2_2_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4220", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4220.Show
'   Me.MousePointer = vbDefault
'End Sub

'科目餘額查詢
'Remove by Lydia 2017/02/24 改成Main4_2_1
'Private Sub Main4_2_3_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4230", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4230.Show
'   Me.MousePointer = vbDefault
'End Sub

'科目明細查詢(對沖)
'Remove by Lydia 2017/02/24 改成Main4_2_1
'Private Sub Main4_2_4_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4240", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4240.Show
'   Me.MousePointer = vbDefault
'End Sub

'智權人員點數查詢
'Remove by Lydia 2017/02/24 改成Main4_2_1
'Private Sub Main4_2_5_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4250", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4250.Show
'   Me.MousePointer = vbDefault
'End Sub

'單據/傳票號碼查詢
'Remove by Lydia 2017/02/24 改成Main4_2_1
'Private Sub Main4_2_6_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4260", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4260.Show
'   Me.MousePointer = vbDefault
'End Sub

'智權人員結餘點數查詢
'Remove by Lydia 2017/02/24 改成Main4_2_1
'Private Sub Main4_2_7_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4270", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4270.Show
'   Me.MousePointer = vbDefault
'End Sub

'Add By Cheng 2003/05/27 日記帳查詢
'Remove by Lydia 2017/02/24 改成Main4_2_1
'Private Sub Main4_2_8_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4280", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4280.Show
'   Me.MousePointer = vbDefault
'End Sub

'簽收資料查詢
'Remove by Lydia 2017/02/24 改成Main4_2_1
'Private Sub Main4_2_9_Click()
'   Toolbar1.Visible = False
'   StatusBar1.Visible = False
'   MenuDisabled
'   'Modified by Lydia 2017/01/26 是否需要輸入密碼
'  ' frm210106.Show
'    If frm210106_1.setNextForm = "" Then
'       frm210106.Show
'    Else
'       frm210106_1.setCaller frm210106
'       frm210106_1.Show
'    End If
'End Sub

Private Sub Main4_3_1_Click(Index As Integer)

   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   Me.MousePointer = vbHourglass
   Select Case Index
      Case 1   '應收/付轉傳票作業
         If CheckUse("Frmacc4350", strExec) = False Then
            Exit Sub
         End If
         'Add by Amy 2022/06/07 判斷與自動轉傳票程式不可同時執行
         If TranNoLock("Frmacc4350") = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc4350.Show
         
      Case 2   '過帳及分攤作業
         If CheckUse("Frmacc4320", strExec) = False Then
            Exit Sub
         End If
         'ADD BY SONIA 2014/8/8 為免過帳,月結算,年度結轉同時開啟,無法檢查每一作業的前一步驟是否完成,造成報表錯誤,故加入此控制
         If PUB_GetLock("Frmacc4320", "", "過帳 或 月結算 或 年度結轉") = False Then
            Exit Sub
         End If
         'END 2014/8/8
         tool3_enabled
         MenuDisabled
         Frmacc4320.Show
      
      Case 3   '月結算作業
         If CheckUse("Frmacc4330", strExec) = False Then
            Exit Sub
         End If
         'ADD BY SONIA 2014/8/8 為免過帳,月結算,年度結轉同時開啟,無法檢查每一作業的前一步驟是否完成,造成報表錯誤,故加入此控制
         If PUB_GetLock("Frmacc4320", "", "過帳 或 月結算 或 年度結轉") = False Then
            Exit Sub
         End If
         'END 2014/8/8
         tool3_enabled
         MenuDisabled
         Frmacc4330.Show
      
      Case 4   '年度結轉作業
         If CheckUse("Frmacc4340", strExec) = False Then
            Exit Sub
         End If
         'ADD BY SONIA 2014/8/8 為免過帳,月結算,年度結轉同時開啟,無法檢查每一作業的前一步驟是否完成,造成報表錯誤,故加入此控制
         If PUB_GetLock("Frmacc4320", "", "過帳 或 月結算 或 年度結轉") = False Then
            Exit Sub
         End If
         'END 2014/8/8
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc4340.Show
         
      Case 5   '歷史傳票資料刪除處理  2014/3/6 改為不顯示
         If CheckUse("Frmacc4360", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         MenuDisabled
         Frmacc4360.Show         

         
      Case 7   '翻譯費轉應付作業
         If CheckUse("Frmacc1310", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         MenuDisabled
         Frmacc1310.Show
         
      Case 8   '已開發票未收款沖帳作業  2014/3/14 add by sonia
         If CheckUse("Frmacc43a0", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         MenuDisabled
         Frmacc43a0.Show
'Mark by Amy 2016/03/03 搬至基本資料
'      Case 9   '每月業績開放/關閉輸入 Add by Amy 2016/01/11
''         If CheckUse("Frmacc43c0", strExec) = False Then
''            Exit Sub
''         End If
'         tool3_enabled
'         MenuDisabled
'         Frmacc43c0.Show
      'Added by Lydia 2016/02/01
      Case 10  '取消過帳或月(年)結
         If CheckUse("Frmacc43d0", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         MenuDisabled
         Frmacc43d0.Show
   End Select
   Me.MousePointer = vbDefault
   
End Sub

'Modified by Lydia 2017/02/24 改成Index
'Private Sub Main4_4_1_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4410", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4410.Show
'   Me.MousePointer = vbDefault
'End Sub
Private Sub Main4_4_1_Click(Index As Integer)
   'Memo by Amy 2020/08/12 加L公司時,財務告知不使用,故隱藏
   '試算表(acc4440)/部門費用統計表(acc44a0)/部門綜合損益(子科目)(acc44b0)/年度部門損益統計表(acc44d0)
   'end 2020/08/12
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   Select Case Index
      Case 1 '會計傳票列印
        If CheckUse("Frmacc44g0", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc44g0.Show

      Case 2 '日計表
        If CheckUse("Frmacc4410", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc4410.Show

      Case 3 '科目明細表(對沖)
        If CheckUse("Frmacc4430", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc4430.Show
      
      Case 4 '科目餘額表
        If CheckUse("Frmacc4420", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc4420.Show

      Case 5 '試算表
        If CheckUse("Frmacc4440", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc4440.Show

      Case 6  '科目分類帳
        If CheckUse("Frmacc4450", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc4450.Show

      'Modify by Amy 2018/03/07 改順序-瑞婷
      Case 7  '智權人員點數明細表
        If CheckUse("Frmacc44j0", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc44j0.Show
      
      Case 8 '專業點數明細表
        If CheckUse("Frmacc44r0", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc44r0.Show
        
    'Modify by Amy 2023/02/18 +ACS 待分潤明細表
    Case 9
        If CheckUse("Frmacc41l1", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc41l1.Show
        
    'Modify by Amy 2021/09/23 專業達成點數表-秘書畫面顯示不同,選項拆開程式調整較少
    Case 10 '專業達成點數表-秘書
        If CheckUse("Frmacc44r0", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc44r0.stState = "SEC"
        Frmacc44r0.Show
        
       Case 11 '預算實績比較表
        If CheckUse("Frmacc4490", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc4490.Show

      Case 12 '綜合損益比較表
        If CheckUse("Frmacc4470", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc4470.Show
  
      Case 13 '年度綜合損益統計表
        If CheckUse("Frmacc44c0", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc44c0.Show
        
      Case 14 '資產負債表
        If CheckUse("Frmacc4480", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc4480.Show
        
      Case 15 '綜合損益表
        If CheckUse("Frmacc4460", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc4460.Show
   
      Case 16 '年度部門綜合損益統計表
        If CheckUse("Frmacc44d0", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc44d0.Show
        
      Case 17 '資產負債比較表
        If CheckUse("Frmacc44e0", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc44e0.Show

      Case 18  '部門費用統計表
        If CheckUse("Frmacc44a0", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc44a0.Show
   
      Case 19 '部門綜合損益表(子科目)
        If CheckUse("Frmacc44b0", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc44b0.Show

      Case 20 '部門綜合損益表
        If CheckUse("Frmacc44h0", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc44h0.Show

      Case 21 '預算資料列印
        If CheckUse("Frmacc44l0", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc44l0.Show

      Case 22  '會計科目代號對照表
        If CheckUse("Frmacc44k0", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc44k0.Show

      Case 23  '費用科目分攤比率表
        If CheckUse("Frmacc44m0", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc44m0.Show
      'end 2018/03/07

      Case 24 '結餘單列印
      'end 2023/02/18
      'end 2021/09/23
        If CheckUse("frm040330", strExec) = False Then
           Exit Sub
        End If
        frm040330.Show
   End Select
   Me.MousePointer = vbDefault
End Sub
'end 2017/02/24

'Remove by Lydia 2017/02/24 改成Main4_4_1
'Private Sub Main4_4_10_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc44a0", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc44a0.Show
'   Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_4_1
'Private Sub Main4_4_11_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc44b0", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc44b0.Show
'   Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_4_1
'Private Sub Main4_4_12_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc44c0", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc44c0.Show
'   Me.MousePointer = vbDefault
'End Sub
'
''Remove by Lydia 2017/02/24 改成Main4_4_1
'Private Sub Main4_4_13_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc44d0", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc44d0.Show
'   Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_4_1
'Private Sub Main4_4_14_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc44e0", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc44e0.Show
'   Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_4_1
'Private Sub Main4_4_16_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc44g0", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc44g0.Show
'   Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_4_1
'Private Sub Main4_4_17_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc44h0", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc44h0.Show
'   Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_4_1
'Private Sub Main4_4_19_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc44j0", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc44j0.Show
'   Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_4_1
'Private Sub Main4_4_2_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4420", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4420.Show
'   Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_4_1
'Private Sub Main4_4_20_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc44k0", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc44k0.Show
'   Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_4_1
'Private Sub Main4_4_21_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc44l0", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc44l0.Show
'   Me.MousePointer = vbDefault
'End Sub

'費用科目分攤比率表
'Remove by Lydia 2017/02/24 改成Main4_4_1
'Private Sub Main4_4_22_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc44m0", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc44m0.Show
'   Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_4_1
'Private Sub Main4_4_23_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc44r0", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc44r0.Show
'   Me.MousePointer = vbDefault
'End Sub

'edit by nickc 2005/11/10 取消舊結餘程式
'Private Sub Main4_4_24_Click()
'   If CheckUse("frm040320", strExec) = False Then
'      Exit Sub
'   End If
'   frm040320.Show
'End Sub
'Remove by Lydia 2017/02/24 改成Main4_4_1
'Private Sub Main4_4_25_Click()
'   If CheckUse("frm040330", strExec) = False Then
'      Exit Sub
'   End If
'   frm040330.Show
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_4_1
'Private Sub Main4_4_3_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4430", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4430.Show
'   Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_4_1
'Private Sub Main4_4_4_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4440", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4440.Show
'   Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_4_1
'Private Sub Main4_4_5_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4450", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4450.Show
'   Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_4_1
'Private Sub Main4_4_6_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4460", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4460.Show
'   Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_4_1
'Private Sub Main4_4_7_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4470", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4470.Show
'   Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_4_1
'Private Sub Main4_4_8_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4480", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4480.Show
'   Me.MousePointer = vbDefault
'End Sub

'Remove by Lydia 2017/02/24 改成Main4_4_1
'Private Sub Main4_4_9_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc4490", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc4490.Show
'   Me.MousePointer = vbDefault
'End Sub

Private Sub Main4_Click()
   Toolbar1.Visible = True
   StatusBar1.Visible = True
End Sub

Private Sub Main5_Click()
   Toolbar1.Visible = False
   StatusBar1.Visible = False
End Sub

Private Sub Main6_1_10_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc2171", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc2171.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main6_1_11_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc2153", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc2153.Show
   Me.MousePointer = vbDefault
End Sub
'Added by Morgan 2019/10/5
'美金請款匯率資料維護
Private Sub Main6_1_12_Click()
   If CheckUse("Frmacc21m0", strExec) = False Then
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc21m0.Show
   Me.MousePointer = vbDefault
End Sub
'Added by Morgan 2019/10/5
'其他幣別請款匯率資料維護
Private Sub Main6_1_13_Click()
   If CheckUse("Frmacc21s0", strExec) = False Then
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc21s0.Show
   Me.MousePointer = vbDefault
End Sub
'Added by Morgan 2023/3/30
'付款日期調整
Private Sub Main6_1_14_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc2172", strExec) = False Then
      Exit Sub
   End If
   tool10_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc2172.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main6_1_2_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc21h0", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc21h0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main6_1_3_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc21i0", strExec) = False Then
      Exit Sub
   End If
   tool8_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc21i0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main6_1_4_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc21k0", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc21k0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main6_1_5_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc2150", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc2150.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main6_1_6_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc21j0", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc21j0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main6_1_7_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc2160", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc2160.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main6_1_8_Click()
   'FC案件不請款確認維護
   If CheckUse("frm040333", strExec) = False Then
      Exit Sub
   End If
   frm040333.Show
End Sub

'Added by Morgan 2015/12/21
Private Sub Main6_1_9_Click()
   '預估結匯匯率資料維護
   If CheckUse("Frmacc21o0", strExec) = False Then
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   'Modified by Morgan 2019/7/10
   '開放財務處可維護
   'tool3_enabled
   tool1_enabled
   'end 2019/7/10
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc21o0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main6_4_1_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc24b0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc24b0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main6_4_3_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc24f0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc24f0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main6_Click()
   Toolbar1.Visible = True
   StatusBar1.Visible = True
End Sub
'Add by Morgan 2005/12/15 切換連線
Private Sub Main7_0_Click()
   If PUB_Connect2DB(True) = False Then
      Unload Me
   End If
End Sub

Private Sub Main7_1_1_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc5100", strExec) = False Then
      Exit Sub
   End If
   'Modify by Amy 2014/02/14 開放使用修改、前後筆功能
   'tool3_enabled
   tool8_enabled
   'end 2014/02/14
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc5100.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main7_1_2_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc5200", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc5200.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main7_1_3_Click()
   frm880013.Show vbModal
End Sub

Private Sub Main7_1_4_Click()
   frm880011.bolAppOnly = True
   frm880011.Show 1
End Sub

Private Sub Main7_1_5_Click()
   frmChgUser.Caption = "解除畫面擷取限制"
   frmChgUser.SSTab1.TabVisible(1) = True
   frmChgUser.SSTab1.TabVisible(0) = False
   frmChgUser.Show
End Sub

Private Sub Main7_2_2_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc6200.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main7_3_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   'Modified by Morgan 2025/7/31
   'End
   Unload Me
   'end 2025/7/31
End Sub

Private Sub Main7_Click()
   Toolbar1.Visible = True
   StatusBar1.Visible = True
End Sub

'Add by Amy 2014/01/17 將原付款作業拆成兩項
Private Sub Main8_1_1_1_Click(Index As Integer)
    If strFormName <> MsgText(601) Then
        Exit Sub
    End If
    Select Case Index
        Case 1
            If CheckUse("Frmacc1180", strExec) = False Then
                Exit Sub
            End If
            tool1_enabled
            Me.MousePointer = vbHourglass
            MenuDisabled
            Frmacc1180.Show
            Me.MousePointer = vbDefault
        Case 2
'            If CheckUse("Frmacc1185", strExec) = False Then
'                Exit Sub
'            End If
'            tool1_enabled
'            Me.MousePointer = vbHourglass
'            MenuDisabled
'            Frmacc1185.Show
'            Me.MousePointer = vbDefault
        End Select
End Sub
'end 2014/01/17

Private Sub Main8_1_1_Click(Index As Integer)
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   
   Select Case Index
      Case 1 '廠商基本資料
         If CheckUse("Frmacc1160", strExec) = False Then
            Exit Sub
         End If
         tool1_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1160.Show
         Me.MousePointer = vbDefault
      
      Case 2 '應付款資料
         If CheckUse("Frmacc1170", strExec) = False Then
            Exit Sub
         End If
         tool1_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1170.Show
         Me.MousePointer = vbDefault
      
      Case 3 '付款作業
         If CheckUse("Frmacc1180", strExec) = False Then
            Exit Sub
         End If
         tool1_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1180.Show
         Me.MousePointer = vbDefault
      
      Case 4 '員工翻譯費率維護
         If CheckUse("Frmacc11j0", strExec) = False Then
            Exit Sub
         End If
         tool1_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc11j0.Show
         Me.MousePointer = vbDefault
      
      Case 5 '翻譯費資料輸入
         If CheckUse("Frmacc11k0", strExec) = False Then
            Exit Sub
         End If
         tool1_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc11k0.Show
         Me.MousePointer = vbDefault
   End Select
End Sub

Private Sub Main8_2_1_Click(Index As Integer)
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   
   Select Case Index
      Case 1 '應付款查詢
         If CheckUse("Frmacc1270", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1270.Show
         Me.MousePointer = vbDefault
      
      Case 2 '客戶回執資料查詢/列印/回收
         If CheckUse("Frmacc12a0", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc12a0.Show
         Me.MousePointer = vbDefault
         
      Case 3 '翻譯費查詢
         If CheckUse("Frmacc12b0", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc12b0.Show
         Me.MousePointer = vbDefault
      
      Case 4 'FCP翻譯完稿案件查詢/列印
         If CheckUse("frm060203", strExec) = False Then
            Exit Sub
         End If
         
         Toolbar1.Visible = False
         StatusBar1.Visible = False
         MenuDisabled
         frm060203.Show
         
      Case 5 '翻譯費用及請款明細查詢/列印
         If CheckUse("frm060208", strExec) = False Then
            Exit Sub
         End If
         
         Toolbar1.Visible = False
         StatusBar1.Visible = False
         MenuDisabled
         frm060208.Show
      'Added by Lydia 2024/07/29
      Case 6 '出庭費查詢 (113/11/01上線)
        If CheckUse("frm075013_2", strExec, False) = False Then
           Exit Sub
        End If
        Toolbar1.Visible = False
        StatusBar1.Visible = False
        MenuDisabled
        frm075013_2.Show
   End Select
End Sub

Private Sub Main8_4_1_Click(Index As Integer)
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   Select Case Index
      Case 1 '付款工作底稿
         If CheckUse("Frmacc1430", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1430.Show
         Me.MousePointer = vbDefault
      
      Case 2 '客戶付款明細 'Memo by Amy 2022/03/14 原:付款明細-瑞婷
         If CheckUse("Frmacc14b0", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc14b0.Show
         Me.MousePointer = vbDefault
         
      Case 3 '國內廠商付款明細表 'Memo by Amy 2022/03/14 原:國內付款明細表-瑞婷
         If CheckUse("Frmacc14d0", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc14d0.Show
         Me.MousePointer = vbDefault
                  
      Case 4 '應付款統計表
         If CheckUse("Frmacc14c0", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc14c0.Show
         Me.MousePointer = vbDefault
         
      Case 5 '名條列印
         If CheckUse("Frmacc14g0", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc14g0.Show
         Me.MousePointer = vbDefault
         
      'Add by Morgan 2006/6/13
      Case 6 '廠商付款明細表
         If CheckUse("Frmacc14i0", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc14i0.Show
         Me.MousePointer = vbDefault
         
      'Add by Morgan 2007/6/4
      Case 7 '翻譯費總表
         If CheckUse("Frmacc14j0", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         frmacc14j0.Show
         Me.MousePointer = vbDefault
         
      'Add by Morgan 2007/6/5
      Case 8 '翻譯費明細表
         If CheckUse("Frmacc14k0", strExec) = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         frmacc14k0.Show
         Me.MousePointer = vbDefault
   End Select
End Sub

Private Sub Main8_Click()
   Toolbar1.Visible = True
   StatusBar1.Visible = True
End Sub

Private Sub Main9_1_1_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc11b0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc11b0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main9_1_2_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc11c0", strExec) = False Then
      Exit Sub
   End If
   'tool3_enabled
   tool13_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc11c0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main9_1_3_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc11f0", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Screen.MousePointer = vbHourglass
   MenuDisabled
   Frmacc11f0.Show
   Screen.MousePointer = vbDefault
End Sub

Private Sub Main9_1_4_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc11g0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc11g0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main9_1_5_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc11h0", strExec) = False Then
      Exit Sub
   End If
   tool7_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc11h0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main9_4_1_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc44i0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc44i0.Show
   Me.MousePointer = vbDefault
End Sub

'Add by Amy 2025/11/14 年度未扣繳查詢
Private Sub Main9_4_11_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc44t1", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc44t1.Show
   Me.MousePointer = vbDefault
End Sub

'CANCEL BY SONIA 2013/7/2 辜說沒有在用了
'Private Sub Main9_4_2_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc44o0", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc44o0.Show
'   Me.MousePointer = vbDefault
'End Sub

Private Sub Main9_4_3_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc44q0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc44q0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main9_4_4_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc44p0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc44p0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main9_4_5_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc11e0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Screen.MousePointer = vbHourglass
   MenuDisabled
   Frmacc11e0.Show
   Screen.MousePointer = vbDefault
End Sub

Private Sub Main9_4_6_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc44t0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc44t0.Show
   Me.MousePointer = vbDefault
End Sub

'Add By Sindy 2016/11/9
Private Sub Main9_4_7_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc44w0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc44w0.Show
   Me.MousePointer = vbDefault
End Sub

'Add By Sindy 2016/11/15
Private Sub Main9_4_8_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc44y0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc44y0.Show
   Me.MousePointer = vbDefault
End Sub

'Add By Sindy 2017/11/3 年度扣繳檢核(抬頭及信箱)
Private Sub Main9_4_10_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc44x0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc44x0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main9_Click()
   Toolbar1.Visible = True
   StatusBar1.Visible = True
End Sub

Private Sub MDIForm_Activate()
Dim strText As Variant
Dim nResponse

'Added by Morgan 2012/1/5 視窗原為最大化,改指定大小及位置以方便其他應用程式操作
Static bolFormIsSet As Boolean
If bolFormIsSet = False Then
   PUB_InitFormPos Me
   bolFormIsSet = True
End If
'end 2012/1/5
   
   'Add By Sindy 2011/9/15 檢查人事出缺勤是否有待辦事件
   '只要執行一次
   If m_blnABSActivated = False Then
      m_blnABSActivated = True
      pub_CallNextABSForm = False
    
      '檢查人事出缺勤是否有待辦事件,若有,顯示訊息提醒操作人員
      strText = ChkIsAbsenceMustProMsg
      If strText <> "" Then
         'Modify By Sindy 2020/5/28
         'nResponse = MsgBox("您有" & strText & "，現在是否要進行處理？", vbYesNo + vbCritical + vbQuestion, "電子簽核")
         MsgBox "您有" & strText & "，請至「查詢系統」處理！", , "電子簽核"
         '2020/5/28 END
'         If nResponse = vbYes Then
'            pub_CallNextABSForm = True
'            strText = ChkIsAbsenceMustPro
'            If InStr(1, strText, "A") > 0 Then
'               frm180201.Show
'            ElseIf InStr(1, strText, "B") > 0 Then
'               frm180101.Show
'            ElseIf InStr(1, strText, "C") > 0 Then
''               frm160201.intChoose = 1
''               frm160201.Hide
''               Call frm160201.cmdOK_Click(0)
'''               Unload frm160201
'               frm180203_1.Show
'            ElseIf InStr(1, strText, "D") > 0 Then
'               frm160102.intChoose = 1
'               frm160102.Hide
'               Call frm160102.cmdok_Click(0)
''               Unload frm160102
'            End If
'         End If
      End If
      
      ChkCFMailSchedule 'Added by Morgan 2024/6/11 檢查是否要新增索取CF對帳單的排程
   End If
End Sub

''Add By Sindy 2011/10/7
'Public Sub SysStartCallForm()
'   '此函數在各系統一啟動時,因出缺勤待辦提示納入之故,共用會使用到,所以不可刪除
'End Sub

'*************************************************
'  工具列按鈕及圖案設定
'
'*************************************************
Private Sub MDIForm_Load()
Dim lngValue, lngBufferSize As Long, intCount As Integer
Dim strUserId As String * 10, strLocalId As String

    If pub_str_LoginSucceeded = "1" Then
'       Show
       
       'Removed by Morgan 2012/1/5 會影響 Activate 事件的觸發
       'Me.Enabled = False
       'Frmacc0002.Show
       'DoEvents
       'Unload Frmacc0002
       'Me.Enabled = True
       'end 2012/1/5
       
       lngBufferSize = 10
        If strUserNum = "" Then
        '   lngValue = WNetGetUser(strLocalId, strUserId, lngBufferSize)
           lngValue = getUserName(strUserId, lngBufferSize)
           For intCount = 1 To 10
              If Asc(Mid(strUserId, intCount, 1)) = 0 Then
                 Exit For
              End If
              strUserNum = strUserNum & Mid(strUserId, intCount, 1)
              'strUserNum = strUserNum 'Remove by Lydia 2017/05/12
           Next intCount
           strUserNum = UCase(strUserNum) 'Added by Lydia 2017/05/12
        End If
       Me.Caption = Me.Caption & "  [ " & strUserNum & " ]"
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
       tool4_enabled
       Main_C
       strFormName = MsgText(601)
       strExitControl = MsgText(602)
       For intCounter = 1 To 4
          StatusBar1.Panels.add
       Next intCounter
       StatusBar1.Height = 300
       StatusBar1.Panels.Item(1).Width = 5500
       StatusBar1.Panels.Item(2).Width = 1000
       StatusBar1.Panels.Item(3).Text = CFDate(ACDate(ServerDate))
       StatusBar1.Panels.Item(4).Text = time
'       Me.Icon = LoadPicture(strIcoPath)
       Timer1.Interval = 1000
       Systemkind_g = GetSystemKindByNick
       Systemkind_g_P = GetSystemKindByNickP
       Systemkind_g_T = GetSystemKindByNickT
       Systemkind_g_TnoS = GetSystemKindByNickTnoS
        
        'Removed by Morgan 2014/11/6 移到 aacc_start
        '取得使用者所別
        'pub_strUserOffice = PUB_GetST06(strUserNum)
        'end 2014/11/6
        
        'Add By Cheng 2004/05/18
        '設定系統日期變數
        strSrvDate(1) = ServerDate
        strSrvDate(2) = ServerDate - 19110000
        'End
    Else
        Me.Visible = False
        If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") = 0 Then
            'Modified by Morgan 2013/5/8 統一用frmLogin
            'frmLogin_1.Show vbModal
            frmLogin.Show vbModal
        Else
            Frmacc0000.Timer2.Interval = 0
            pub_str_LoginSucceeded = "1"
            MDIForm_Load
        End If
    End If
    If strUserName = "" Then strUserName = GetPrjSalesNM(strUserNum)  'Add by Morgan 2006/12/11

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   PUB_AddAuditLog AL_登出 'Added by Morgan 2025/7/31
   EndOfficeAp 'Added by Morgan 2025/9/4
'   objOraDatabase.Close
'   objOraSession.Close
   adoTaie.Close
   Set Frmacc0000 = Nothing
End Sub


'*************************************************
'  狀態列時間設定
'
'*************************************************
Private Sub Timer1_Timer()
   StatusBar1.Panels.Item(4).Text = time
End Sub

Private Sub Timer2_Timer()
    '若登入失敗
    If pub_str_LoginSucceeded <> "1" Then
        Me.Timer1.Interval = 0
    '若登入成功
    Else
'        Me.Timer1.Interval = 100
        Me.Timer2.Interval = 0
'        MDIForm_Load
'        Me.Show
    End If
End Sub

'Add by Morgan 2005/3/2 控制不可拷貝畫面
Private Sub Timer3_Timer()
   Static dtNow As Date 'Added by Morgan 2024/8/7
      
On Error Resume Next 'Added by Morgan 2019/10/31 若有其他軟體也在使用剪貼簿時會發生521(無法開啟剪貼簿)的錯誤(Ex.Word開啟剪貼簿並擷取畫面)

   'Added by Morgan 2024/8/7 定時執行一次語法以確保跨網段連線時網路不會被切斷
   If Now > dtNow Then
      dtNow = DateAdd("n", cntAutoQueryInterval, Now)
      ClsLawReadRstMsg 1, "select * from dual"
   End If
   'end 2024/8/7

'add by nickc 2005/05/02 電腦中心的不管
If Pub_StrUserSt03 = "M51" Or Pub_Can_Copy_Pic = True Then Exit Sub
    '圖檔才清
    If Clipboard.GetFormat(1) = False And Clipboard.GetFormat(2) = True And Clipboard.GetFormat(3) = False Then

        Clipboard.Clear
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

'Added by Lydia 2016/11/7
'國外固定寄催款單代理人維護
Private Sub Main2_1_22_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc21w0", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc21w0.Show
   Me.MousePointer = vbDefault
End Sub

'Added by Lydia 2016/12/19 會計師客戶資料查詢
Private Sub Main9_4_9_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc44z0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc44z0.Show
   Me.MousePointer = vbDefault
End Sub

'Added by Lydia 2017/02/24  財產目錄
Private Sub Main4_5_1_Click(Index As Integer)
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   Select Case Index
      Case 1 '財產目錄作業
        If CheckUse("Frmacc41i0", strExec) = False Then
           Exit Sub
        End If
        tool1_enabled
        MenuDisabled
        Frmacc41i0.Show
      Case 2 '財產報廢作業
        If CheckUse("Frmacc41i0_1", strExec) = False Then
           Exit Sub
        End If
        'Modified by Lydia 2017/05/15 取消刪除功能
        'tool1_enabled
        tool14_enabled
        'end 2017/05/15
        MenuDisabled
        Frmacc41i0_1.Show
      Case 3 '財產目錄表
        If CheckUse("Frmacc4510", strExec) = False Then
           Exit Sub
        End If
        tool3_enabled
        MenuDisabled
        Frmacc4510.Show
   End Select
   Me.MousePointer = vbDefault

End Sub

'Added by Lydia 2019/08/16 客戶應收帳款收文檢查上限
'Mark by Lydia 2020/02/07 開放維護功能
'Private Sub Main1_2_15_Click()
'    If CheckUse("frm140502", strExec) = True Then
'       frm140502.Show
'    End If
'End Sub

'Added by Lydia 2019/08/16 客戶特殊付款週期維護
'Mark by Lydia 2020/02/07 開放維護功能
'Private Sub Main1_2_16_Click()
'    If CheckUse("frm140504", strExec) = True Then
'       frm140504.Show
'    End If
'End Sub

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

'Added by Morgan 2024/6/11
'檢查是否需新增索取CF對帳單排程(每年6月)
Private Sub ChkCFMailSchedule()
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   Dim srtMsg As String, bolMail1 As Boolean, bolMail2 As Boolean
   
   'Modified by Morgan 2025/6/11 改每年7月
   If Val(Mid(strSrvDate(1), 5, 2)) = 7 Then
      If Pub_GetSpecMan("外專請款單已收款通知人員") = strUserNum Then
         If ServerTime < "170000" Then '發信時間固定設18:00,控制17:00以前執行這樣才有緩衝時間
            stSQL = "select * from mailschedule where ms15=1024 and ms08>=" & Left(strSrvDate(1), 4) & "0101"
            intR = 1
            Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
            If intR = 0 Then
               bolMail1 = True
               srtMsg = "索取CF對帳單(中文)"
            End If
            
            stSQL = "select * from mailschedule where ms15=2048 and ms08>=" & Left(strSrvDate(1), 4) & "0101"
            intR = 1
            Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
            If intR = 0 Then
               bolMail2 = True
               If bolMail1 Then
                  srtMsg = "索取CF對帳單(中文及英文)"
               Else
                  srtMsg = "索取CF對帳單(英文)"
               End If
            End If
            
            If bolMail1 Or bolMail2 Then
               If MsgBox("現在是否要執行新增「" & srtMsg & "」排程？", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                  frm140410_1.m_AutoRun = True
                  frm140410_1.m_Schedule1 = bolMail1
                  frm140410_1.m_Schedule2 = bolMail2
                  frm140410_1.Show
               End If
            End If
         End If
      End If
   End If
   Set rsQuery = Nothing
End Sub

'Added by Lydia 2024/07/29
'以名稱取得表單--通用不可刪
Public Function GetForm(pFormName As String) As Form
   Select Case pFormName
      'Add By Sindy 2024/11/5
      Case "frm090801_Q"
         Set GetForm = frm090801_Q
         '2024/11/5 END
   End Select
End Function
