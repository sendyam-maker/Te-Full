VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H80000018&
   Caption         =   "專利及承辦人系統"
   ClientHeight    =   5484
   ClientLeft      =   960
   ClientTop       =   2712
   ClientWidth     =   10440
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  '最大化
   Begin VB.Timer tmrSalary 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   630
      Top             =   3240
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3780
      Top             =   2610
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5010
      Top             =   2520
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   720
      Top             =   1230
   End
   Begin VB.Timer tmrConnect 
      Left            =   735
      Top             =   1710
   End
   Begin VB.Timer Timer2 
      Left            =   30
      Top             =   1650
   End
   Begin VB.Timer Timer1 
      Left            =   30
      Top             =   1230
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '對齊表單下方
      Height          =   276
      Left            =   0
      TabIndex        =   1
      Top             =   5208
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   487
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
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   1016
      ButtonWidth     =   487
      ButtonHeight    =   889
      Appearance      =   1
      _Version        =   393216
      Begin VB.ListBox List1 
         Height          =   228
         Left            =   1980
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox TextComp 
         Height          =   285
         Left            =   0
         TabIndex        =   2
         Text            =   "暫存文字框"
         Top             =   0
         Visible         =   0   'False
         Width           =   1905
      End
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   600
      Top             =   612
      _ExtentX        =   974
      _ExtentY        =   974
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   1176
      Top             =   600
      _ExtentX        =   974
      _ExtentY        =   974
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   -1  'True
      NewSession      =   0   'False
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
      Caption         =   "專利"
      Index           =   3
      Begin VB.Menu mnuTitle2 
         Caption         =   "內專"
         Index           =   1
         Begin VB.Menu mnu04 
            Caption         =   "資料處理"
            Index           =   1
            Begin VB.Menu mnu0401 
               Caption         =   "分案"
               Index           =   1
            End
            Begin VB.Menu mnu0401 
               Caption         =   "各式申請書／書表"
               Index           =   2
            End
            Begin VB.Menu mnu0401 
               Caption         =   "國外指示信"
               Index           =   3
            End
            Begin VB.Menu mnu0401 
               Caption         =   "發文"
               Index           =   4
            End
            Begin VB.Menu mnu0401 
               Caption         =   "申請案號輸入"
               Index           =   5
            End
            Begin VB.Menu mnu0401 
               Caption         =   "審查機關來函"
               Index           =   6
               Begin VB.Menu mnu040105 
                  Caption         =   "實審通知日輸入"
                  Index           =   1
               End
               Begin VB.Menu mnu040105 
                  Caption         =   "初審及公佈通知來函輸入"
                  Index           =   2
               End
               Begin VB.Menu mnu040105 
                  Caption         =   "核准函輸入"
                  Index           =   3
               End
               Begin VB.Menu mnu040105 
                  Caption         =   "核駁函輸入"
                  Index           =   4
               End
               Begin VB.Menu mnu040105 
                  Caption         =   "消滅函／視為撤回輸入"
                  Index           =   6
               End
               Begin VB.Menu mnu040105 
                  Caption         =   "一般來函輸入"
                  Index           =   7
               End
               Begin VB.Menu mnu040105 
                  Caption         =   "證書號數輸入"
                  Index           =   8
               End
               Begin VB.Menu mnu040105 
                  Caption         =   "異議/舉發受理函輸入"
                  Index           =   9
               End
               Begin VB.Menu mnu040105 
                  Caption         =   "年費逾期補繳通知函輸入"
                  Index           =   10
               End
               Begin VB.Menu mnu040105 
                  Caption         =   "實審請求期限屆滿前通知函輸入"
                  Index           =   11
               End
            End
            Begin VB.Menu mnu0401 
               Caption         =   "代理人來函"
               Index           =   7
               Begin VB.Menu mnu040106 
                  Caption         =   "已收達/已提申"
                  Index           =   1
               End
               Begin VB.Menu mnu040106 
                  Caption         =   "通知修正"
                  Index           =   2
               End
               Begin VB.Menu mnu040106 
                  Caption         =   "所外鑑定報告結果"
                  Index           =   3
               End
               Begin VB.Menu mnu040106 
                  Caption         =   "其他來函"
                  Index           =   4
               End
               Begin VB.Menu mnu040106 
                  Caption         =   "代理人信件收達管制"
                  Index           =   5
               End
               Begin VB.Menu mnu040106 
                  Caption         =   "已順稿"
                  Index           =   6
               End
            End
            Begin VB.Menu mnu0401 
               Caption         =   "關聯案件資料維護"
               Index           =   8
               Begin VB.Menu mnu040108 
                  Caption         =   "國內外案件資料維護"
                  Index           =   1
               End
               Begin VB.Menu mnu040108 
                  Caption         =   "大陸發明案資料維護"
                  Index           =   2
               End
               Begin VB.Menu mnu040108 
                  Caption         =   "一案兩申請案件資料維護"
                  Index           =   3
               End
               Begin VB.Menu mnu040108 
                  Caption         =   "擬制喪失新穎性案件資料維護"
                  Index           =   4
               End
               Begin VB.Menu mnu040108 
                  Caption         =   "大陸香港案資料維護"
                  Index           =   5
               End
               Begin VB.Menu mnu040108 
                  Caption         =   "大陸澳門案資料維護"
                  Index           =   6
               End
            End
            Begin VB.Menu mnu0401 
               Caption         =   "領證年費整批發文"
               Index           =   13
            End
            Begin VB.Menu mnu0401 
               Caption         =   "待送件區"
               Index           =   15
            End
            Begin VB.Menu mnu0401 
               Caption         =   "電子送件電子檔整批匯入"
               Index           =   16
            End
            Begin VB.Menu mnu0401 
               Caption         =   "電子公文來函"
               Index           =   17
            End
            Begin VB.Menu mnu0401 
               Caption         =   "公文來函文檔整批匯入"
               Index           =   18
            End
            Begin VB.Menu mnu0401 
               Caption         =   "非台灣案發文後暫緩"
               Index           =   19
            End
            Begin VB.Menu mnu0401 
               Caption         =   "收據/回執整批匯入"
               Index           =   20
            End
            Begin VB.Menu mnu0401 
               Caption         =   "代理人來函匯入"
               Index           =   21
            End
            Begin VB.Menu mnu0401 
               Caption         =   "電子收據匯入"
               Index           =   22
            End
            Begin VB.Menu mnu0401 
               Caption         =   "待處理區"
               Index           =   23
            End
            Begin VB.Menu mnu0401 
               Caption         =   "客戶提供文件處理"
               Index           =   24
            End
         End
         Begin VB.Menu mnu04 
            Caption         =   "查詢作業"
            Index           =   2
            Begin VB.Menu mnu0402 
               Caption         =   "代理人新案案件統計"
               Index           =   1
            End
            Begin VB.Menu mnu0402 
               Caption         =   "未請款明細查詢"
               Index           =   3
            End
            Begin VB.Menu mnu0402 
               Caption         =   "審查委員准駁統計"
               Index           =   4
            End
            Begin VB.Menu mnu0402 
               Caption         =   "FC收款請款點數查詢"
               Index           =   5
            End
            Begin VB.Menu mnu0402 
               Caption         =   "代理人案件性質統計"
               Index           =   6
            End
            Begin VB.Menu mnu0402 
               Caption         =   "員工查詢印表記錄資料查詢"
               Index           =   7
            End
            Begin VB.Menu mnu0402 
               Caption         =   "未發文案件查詢"
               Index           =   8
            End
            Begin VB.Menu mnu0402 
               Caption         =   "電子收文接洽單查詢"
               Index           =   9
            End
         End
         Begin VB.Menu mnu04 
            Caption         =   "報表列印"
            Index           =   3
            Begin VB.Menu mnu0403 
               Caption         =   "通知函"
               Index           =   1
               Begin VB.Menu mnu040301 
                  Caption         =   "公開通知函"
                  Index           =   1
               End
               Begin VB.Menu mnu040301 
                  Caption         =   "公告通知函"
                  Index           =   2
               End
               Begin VB.Menu mnu040301 
                  Caption         =   "公告期滿通知函"
                  Index           =   3
               End
               Begin VB.Menu mnu040301 
                  Caption         =   "繳年費/實體審查通知函"
                  Index           =   4
               End
               Begin VB.Menu mnu040301 
                  Caption         =   "其他通知函/聯絡單"
                  Index           =   6
               End
               Begin VB.Menu mnu040301 
                  Caption         =   "年費逾期補繳通知函(整批)"
                  Index           =   7
               End
               Begin VB.Menu mnu040301 
                  Caption         =   "期限通知檢核及報表"
                  Index           =   9
               End
               Begin VB.Menu mnu040301 
                  Caption         =   "未收文期限提醒E-Mail"
                  Index           =   10
               End
            End
            Begin VB.Menu mnu0403 
               Caption         =   "期限管制表"
               Index           =   8
            End
            Begin VB.Menu mnu0403 
               Caption         =   "代理人案件收達/提申管制表"
               Index           =   9
            End
            Begin VB.Menu mnu0403 
               Caption         =   "收文未發文明細表"
               Index           =   10
            End
            Begin VB.Menu mnu0403 
               Caption         =   "催審函/催審表"
               Index           =   11
            End
            Begin VB.Menu mnu0403 
               Caption         =   "智權人員收文明細表"
               Index           =   12
            End
            Begin VB.Menu mnu0403 
               Caption         =   "收文簿"
               Index           =   13
            End
            Begin VB.Menu mnu0403 
               Caption         =   "發文簿"
               Index           =   14
            End
            Begin VB.Menu mnu0403 
               Caption         =   "核准(駁)簿"
               Index           =   15
            End
            Begin VB.Menu mnu0403 
               Caption         =   "大陸發明案參考資料表"
               Index           =   16
            End
            Begin VB.Menu mnu0403 
               Caption         =   "顧問客戶委辦案件明細表"
               Index           =   17
            End
            Begin VB.Menu mnu0403 
               Caption         =   "後金案件表"
               Index           =   18
            End
            Begin VB.Menu mnu0403 
               Caption         =   "延期明細表"
               Index           =   19
            End
            Begin VB.Menu mnu0403 
               Caption         =   "不出名案件明細表"
               Index           =   20
            End
            Begin VB.Menu mnu0403 
               Caption         =   "代理人案件總簿"
               Index           =   21
            End
            Begin VB.Menu mnu0403 
               Caption         =   "客戶案件總簿輸出"
               Index           =   22
            End
            Begin VB.Menu mnu0403 
               Caption         =   "代理人/申請人名單"
               Index           =   23
            End
            Begin VB.Menu mnu0403 
               Caption         =   "地址條列印"
               Index           =   25
            End
            Begin VB.Menu mnu0403 
               Caption         =   "核准領證期限表"
               Index           =   26
            End
            Begin VB.Menu mnu0403 
               Caption         =   "智慧局年費通知核對清單"
               Index           =   27
            End
            Begin VB.Menu mnu0403 
               Caption         =   "證書PDF列印"
               Index           =   28
            End
            Begin VB.Menu mnu0403 
               Caption         =   "資策會專利案件季報表"
               Index           =   29
            End
            Begin VB.Menu mnu0403 
               Caption         =   "資策會收到證書清單"
               Index           =   30
            End
         End
         Begin VB.Menu mnu04 
            Caption         =   "統計報表"
            Index           =   4
            Begin VB.Menu mnu0404 
               Caption         =   "收文統計表"
               Index           =   1
            End
            Begin VB.Menu mnu0404 
               Caption         =   "發文統計表"
               Index           =   2
            End
            Begin VB.Menu mnu0404 
               Caption         =   "准駁預估統計表"
               Index           =   3
            End
            Begin VB.Menu mnu0404 
               Caption         =   "准駁統計總表"
               Index           =   4
            End
            Begin VB.Menu mnu0404 
               Caption         =   "准駁統計明細表"
               Index           =   5
            End
            Begin VB.Menu mnu0404 
               Caption         =   "代理人新案案件統計表"
               Index           =   6
            End
            Begin VB.Menu mnu0404 
               Caption         =   "代理人新案案件年度統計表"
               Index           =   7
            End
            Begin VB.Menu mnu0404 
               Caption         =   "代理人/申請人新申請案排行榜"
               Index           =   8
            End
            Begin VB.Menu mnu0404 
               Caption         =   "逾期未結案統計表"
               Index           =   9
            End
         End
         Begin VB.Menu mnu04 
            Caption         =   "檔案維護"
            Index           =   5
            Begin VB.Menu mnu0405 
               Caption         =   "專利案件基本資料維護"
               Index           =   1
            End
            Begin VB.Menu mnu0405 
               Caption         =   "服務業務基本資料維護"
               Index           =   2
            End
            Begin VB.Menu mnu0405 
               Caption         =   "案件進度資料維護"
               Index           =   3
            End
            Begin VB.Menu mnu0405 
               Caption         =   "下一程序資料維護"
               Index           =   4
            End
            Begin VB.Menu mnu0405 
               Caption         =   "國外代理人資料維護"
               Index           =   5
            End
            Begin VB.Menu mnu0405 
               Caption         =   "變更事項資料維護"
               Index           =   6
            End
            Begin VB.Menu mnu0405 
               Caption         =   "延期記錄資料維護"
               Index           =   7
            End
            Begin VB.Menu mnu0405 
               Caption         =   "案件國家收費表維護"
               Index           =   8
            End
            Begin VB.Menu mnu0405 
               Caption         =   "客戶發明人資料維護"
               Index           =   9
            End
            Begin VB.Menu mnu0405 
               Caption         =   "代理人變更名稱作業"
               Index           =   11
            End
            Begin VB.Menu mnu0405 
               Caption         =   "客戶減免身份維護"
               Index           =   12
            End
            Begin VB.Menu mnu0405 
               Caption         =   "系統特殊設定"
               Index           =   13
            End
            Begin VB.Menu mnu0405 
               Caption         =   "領證報價資料維護"
               Index           =   14
            End
            Begin VB.Menu mnu0405 
               Caption         =   "年費報價資料維護"
               Index           =   15
            End
            Begin VB.Menu mnu0405 
               Caption         =   "依案件性質設定各國催審提申期限"
               Index           =   16
            End
            Begin VB.Menu mnu0405 
               Caption         =   "電子報排程維護"
               Index           =   17
            End
            Begin VB.Menu mnu0405 
               Caption         =   "台灣專利總委任書正本案號維護"
               Index           =   18
            End
            Begin VB.Menu mnu0405 
               Caption         =   "更換FC代理人作業"
               Index           =   19
            End
            Begin VB.Menu mnu0405 
               Caption         =   "申請人指定國外代理人維護"
               Index           =   20
            End
            Begin VB.Menu mnu0405 
               Caption         =   "程序人員核判表維護"
               Index           =   21
            End
         End
         Begin VB.Menu mnu04 
            Caption         =   "專利公報"
            Index           =   6
            Begin VB.Menu mnu0406 
               Caption         =   "國內公報"
               Index           =   1
               Begin VB.Menu mnu040601 
                  Caption         =   "專利公報輸入"
                  Index           =   1
               End
               Begin VB.Menu mnu040601 
                  Caption         =   "市場佔有率統計表"
                  Index           =   2
               End
               Begin VB.Menu mnu040601 
                  Caption         =   "專利公報查詢列印"
                  Index           =   3
               End
               Begin VB.Menu mnu040601 
                  Caption         =   "國內公報代理人名稱查詢"
                  Index           =   4
               End
               Begin VB.Menu mnu040601 
                  Caption         =   "國內市場佔有率查詢"
                  Index           =   5
               End
               Begin VB.Menu mnu040601 
                  Caption         =   "國內公報代理人資料維護"
                  Index           =   6
               End
               Begin VB.Menu mnu040601 
                  Caption         =   "國內公報代理人換事務所作業"
                  Index           =   7
               End
               Begin VB.Menu mnu040601 
                  Caption         =   "國內公報代理人資料列印"
                  Index           =   8
               End
               Begin VB.Menu mnu040601 
                  Caption         =   "專利公報轉檔作業"
                  Index           =   9
               End
               Begin VB.Menu mnu040601 
                  Caption         =   "公報特殊字對照檔"
                  Index           =   10
               End
            End
            Begin VB.Menu mnu0406 
               Caption         =   "大陸公報"
               Index           =   2
               Begin VB.Menu mnu040602 
                  Caption         =   "專利公報輸入"
                  Index           =   1
               End
               Begin VB.Menu mnu040602 
                  Caption         =   "市場佔有率統計表"
                  Index           =   2
               End
               Begin VB.Menu mnu040602 
                  Caption         =   "專利公報查詢列印"
                  Index           =   3
               End
               Begin VB.Menu mnu040602 
                  Caption         =   "開拓函列印"
                  Index           =   4
               End
               Begin VB.Menu mnu040602 
                  Caption         =   "大陸事務所資料維護"
                  Index           =   5
               End
               Begin VB.Menu mnu040602 
                  Caption         =   "開拓客戶資料維護"
                  Index           =   6
               End
            End
            Begin VB.Menu mnu0406 
               Caption         =   "國內公開公報"
               Index           =   3
               Begin VB.Menu mnu040603 
                  Caption         =   "國內公開公報輸入"
                  Index           =   1
               End
               Begin VB.Menu mnu040603 
                  Caption         =   "國內公開後實審輸入"
                  Index           =   2
               End
               Begin VB.Menu mnu040603 
                  Caption         =   "公開市場佔有率統計表"
                  Index           =   3
               End
               Begin VB.Menu mnu040603 
                  Caption         =   "專利公開公報查詢列印"
                  Index           =   4
               End
               Begin VB.Menu mnu040603 
                  Caption         =   "國內公開市場佔有率查詢"
                  Index           =   5
               End
               Begin VB.Menu mnu040603 
                  Caption         =   "專利公開公報轉檔作業"
                  Index           =   6
               End
            End
            Begin VB.Menu mnu0406 
               Caption         =   "公開及公告市場統計表"
               Index           =   4
            End
            Begin VB.Menu mnu0406 
               Caption         =   "公報產業分類案件市佔分析"
               Index           =   5
            End
         End
         Begin VB.Menu mnu04 
            Caption         =   "專利公報excel"
            Index           =   7
            Begin VB.Menu mnu0407 
               Caption         =   "專利公報市場排名"
               Index           =   1
            End
            Begin VB.Menu mnu0407 
               Caption         =   "專利公報市場占有率比較"
               Index           =   2
            End
            Begin VB.Menu mnu0407 
               Caption         =   "各單位專利公報件數統計"
               Index           =   3
            End
            Begin VB.Menu mnu0407 
               Caption         =   "專利公報國內各區同業排名"
               Index           =   4
            End
            Begin VB.Menu mnu0407 
               Caption         =   "專利公報國外同業排名"
               Index           =   5
            End
            Begin VB.Menu mnu0407 
               Caption         =   "國籍及洲別統計(含同業)"
               Index           =   6
            End
         End
      End
      Begin VB.Menu mnuTitle2 
         Caption         =   "CFP"
         Index           =   2
         Begin VB.Menu mnu05 
            Caption         =   "資料處理"
            Index           =   1
            Begin VB.Menu mnu0501 
               Caption         =   "分案"
               Index           =   1
            End
            Begin VB.Menu mnu0501 
               Caption         =   "發文"
               Index           =   2
            End
            Begin VB.Menu mnu0501 
               Caption         =   "代理人案件提申"
               Index           =   3
            End
            Begin VB.Menu mnu0501 
               Caption         =   "審查機關來函"
               Index           =   4
               Begin VB.Menu mnu050104 
                  Caption         =   "一般來函輸入"
                  Index           =   1
               End
               Begin VB.Menu mnu050104 
                  Caption         =   "公開/公告資料輸入"
                  Index           =   2
               End
               Begin VB.Menu mnu050104 
                  Caption         =   "證書號數輸入"
                  Index           =   3
               End
               Begin VB.Menu mnu050104 
                  Caption         =   "消滅函輸入"
                  Index           =   4
               End
               Begin VB.Menu mnu050104 
                  Caption         =   "年費逾期補繳通知函"
                  Index           =   5
               End
               Begin VB.Menu mnu050104 
                  Caption         =   "實體審查、領證費逾期補繳通知函"
                  Index           =   6
               End
            End
            Begin VB.Menu mnu0501 
               Caption         =   "代理人來函"
               Index           =   5
               Begin VB.Menu mnu050105 
                  Caption         =   "已收達"
                  Index           =   1
               End
               Begin VB.Menu mnu050105 
                  Caption         =   "通知修正"
                  Index           =   2
               End
               Begin VB.Menu mnu050105 
                  Caption         =   "其他來函"
                  Index           =   3
               End
               Begin VB.Menu mnu050105 
                  Caption         =   "代理人信件收達管制"
                  Index           =   4
               End
            End
            Begin VB.Menu mnu0501 
               Caption         =   "國內外案件資料維護"
               Index           =   6
            End
            Begin VB.Menu mnu0501 
               Caption         =   "美國IDS資料對照維護"
               Index           =   7
            End
            Begin VB.Menu mnu0501 
               Caption         =   "國內外案件資料刪除作業"
               Index           =   8
            End
            Begin VB.Menu mnu0501 
               Caption         =   "一案兩申請案件資料維護"
               Index           =   9
            End
            Begin VB.Menu mnu0501 
               Caption         =   "CFP申請文件齊備維護"
               Index           =   10
            End
            Begin VB.Menu mnu0501 
               Caption         =   "待送件區"
               Index           =   11
            End
            Begin VB.Menu mnu0501 
               Caption         =   "待處理區"
               Index           =   12
            End
            Begin VB.Menu mnu0501 
               Caption         =   "代理人來函匯入"
               Index           =   13
            End
            Begin VB.Menu mnu0501 
               Caption         =   "指示信判發作業"
               Index           =   14
            End
            Begin VB.Menu mnu0501 
               Caption         =   "外翻人員給案維護"
               Index           =   15
            End
         End
         Begin VB.Menu mnu05 
            Caption         =   "查詢作業"
            Index           =   2
            Begin VB.Menu mnu0502 
               Caption         =   "代理人新案案件統計"
               Index           =   1
            End
            Begin VB.Menu mnu0502 
               Caption         =   "未請款明細查詢"
               Index           =   3
            End
            Begin VB.Menu mnu0502 
               Caption         =   "代理人案件性質統計"
               Index           =   4
            End
            Begin VB.Menu mnu0502 
               Caption         =   "互惠代理人目標給案未輸入明細表"
               Index           =   6
            End
            Begin VB.Menu mnu0502 
               Caption         =   "未發文案件查詢"
               Index           =   7
            End
            Begin VB.Menu mnu0502 
               Caption         =   "CF代理人報價附件查詢"
               Index           =   8
            End
         End
         Begin VB.Menu mnu05 
            Caption         =   "報表列印"
            Index           =   3
            Begin VB.Menu mnu0503 
               Caption         =   "詢問進度函"
               Index           =   1
            End
            Begin VB.Menu mnu0503 
               Caption         =   "期限管制表"
               Index           =   2
            End
            Begin VB.Menu mnu0503 
               Caption         =   "代理人案件收達/提申管制表"
               Index           =   3
            End
            Begin VB.Menu mnu0503 
               Caption         =   "收文未發文明細表"
               Index           =   4
            End
            Begin VB.Menu mnu0503 
               Caption         =   "催審表"
               Index           =   5
            End
            Begin VB.Menu mnu0503 
               Caption         =   "業務員案件明細表"
               Index           =   6
            End
            Begin VB.Menu mnu0503 
               Caption         =   "收文簿"
               Index           =   7
            End
            Begin VB.Menu mnu0503 
               Caption         =   "收文明細表"
               Index           =   8
            End
            Begin VB.Menu mnu0503 
               Caption         =   "承辦人發文明細表"
               Index           =   9
            End
            Begin VB.Menu mnu0503 
               Caption         =   "發文點數明細表"
               Index           =   10
            End
            Begin VB.Menu mnu0503 
               Caption         =   "新案承辦人明細表"
               Index           =   11
            End
            Begin VB.Menu mnu0503 
               Caption         =   "承辦人准駁明細表"
               Index           =   12
            End
            Begin VB.Menu mnu0503 
               Caption         =   "期限通知管制表"
               Index           =   13
            End
            Begin VB.Menu mnu0503 
               Caption         =   "取消收文明細表"
               Index           =   14
            End
            Begin VB.Menu mnu0503 
               Caption         =   "後金案件表"
               Index           =   15
            End
            Begin VB.Menu mnu0503 
               Caption         =   "延期明細表"
               Index           =   16
            End
            Begin VB.Menu mnu0503 
               Caption         =   "代理人案件總簿"
               Index           =   17
            End
            Begin VB.Menu mnu0503 
               Caption         =   "客戶案件總簿輸出"
               Index           =   18
            End
            Begin VB.Menu mnu0503 
               Caption         =   "代理人/申請人名單"
               Index           =   19
            End
            Begin VB.Menu mnu0503 
               Caption         =   "地址條列印"
               Index           =   21
            End
            Begin VB.Menu mnu0503 
               Caption         =   "TNT列印"
               Index           =   22
               Visible         =   0   'False
            End
            Begin VB.Menu mnu0503 
               Caption         =   "DHL列印"
               Index           =   23
            End
            Begin VB.Menu mnu0503 
               Caption         =   "美國發明退公開費報表/指示信"
               Index           =   24
            End
            Begin VB.Menu mnu0503 
               Caption         =   "未收文期限提醒E-Mail"
               Index           =   25
            End
            Begin VB.Menu mnu0503 
               Caption         =   "期限通知檢核及報表列印"
               Index           =   26
            End
         End
         Begin VB.Menu mnu05 
            Caption         =   "統計報表"
            Index           =   4
            Begin VB.Menu mnu0504 
               Caption         =   "承辦人收文統計表"
               Index           =   1
            End
            Begin VB.Menu mnu0504 
               Caption         =   "承辦人發文統計表"
               Index           =   2
            End
            Begin VB.Menu mnu0504 
               Caption         =   "准駁統計表"
               Index           =   3
            End
            Begin VB.Menu mnu0504 
               Caption         =   "代理人新案案件統計表"
               Index           =   4
            End
            Begin VB.Menu mnu0504 
               Caption         =   "代理人新案案件年度統計表"
               Index           =   5
            End
            Begin VB.Menu mnu0504 
               Caption         =   "代理人/申請人新申請案排行榜"
               Index           =   6
            End
            Begin VB.Menu mnu0504 
               Caption         =   "逾期未結案統計表"
               Index           =   7
            End
            Begin VB.Menu mnu0504 
               Caption         =   "互惠代理人案件統計表"
               Index           =   8
            End
         End
         Begin VB.Menu mnu05 
            Caption         =   "檔案維護"
            Index           =   5
            Begin VB.Menu mnu0505 
               Caption         =   "專利案件基本資料維護"
               Index           =   1
            End
            Begin VB.Menu mnu0505 
               Caption         =   "服務業務基本資料維護"
               Index           =   2
            End
            Begin VB.Menu mnu0505 
               Caption         =   "案件進度資料維護"
               Index           =   3
            End
            Begin VB.Menu mnu0505 
               Caption         =   "下一程序資料維護"
               Index           =   4
            End
            Begin VB.Menu mnu0505 
               Caption         =   "國外代理人資料維護"
               Index           =   5
            End
            Begin VB.Menu mnu0505 
               Caption         =   "變更事項資料維護"
               Index           =   6
            End
            Begin VB.Menu mnu0505 
               Caption         =   "延期記錄資料維護"
               Index           =   7
            End
            Begin VB.Menu mnu0505 
               Caption         =   "案件國家收費表維護"
               Index           =   8
            End
            Begin VB.Menu mnu0505 
               Caption         =   "客戶發明人資料維護"
               Index           =   9
            End
            Begin VB.Menu mnu0505 
               Caption         =   "代理人變更名稱作業"
               Index           =   11
            End
            Begin VB.Menu mnu0505 
               Caption         =   "申請人國外ID資料維護"
               Index           =   12
            End
            Begin VB.Menu mnu0505 
               Caption         =   "客戶資料維護"
               Index           =   13
            End
            Begin VB.Menu mnu0505 
               Caption         =   "客戶減免身份維護"
               Index           =   14
            End
            Begin VB.Menu mnu0505 
               Caption         =   "特殊人員設定"
               Index           =   15
            End
            Begin VB.Menu mnu0505 
               Caption         =   "國外代理人目標給案量維護"
               Index           =   16
            End
            Begin VB.Menu mnu0505 
               Caption         =   "CFP領證報價資料維護"
               Index           =   17
            End
            Begin VB.Menu mnu0505 
               Caption         =   "非本所實質客戶資料維護"
               Index           =   18
               Visible         =   0   'False
            End
            Begin VB.Menu mnu0505 
               Caption         =   "CF代理人報價附件維護"
               Index           =   19
            End
            Begin VB.Menu mnu0505 
               Caption         =   "申請人指定國外代理人維護"
               Index           =   20
            End
            Begin VB.Menu mnu0505 
               Caption         =   "CFP核駁報價資料維護"
               Index           =   21
            End
            Begin VB.Menu mnu0505 
               Caption         =   "CFP維持費/延展費資料維護"
               Index           =   22
            End
         End
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "程序"
      Index           =   9
      Begin VB.Menu mnu09 
         Caption         =   "撰寫信函作業"
         Index           =   4
      End
      Begin VB.Menu mnu09 
         Caption         =   "P案國外新案指示信"
         Index           =   5
      End
      Begin VB.Menu mnu09 
         Caption         =   "P案各式申請書"
         Index           =   6
      End
      Begin VB.Menu mnu09 
         Caption         =   "主管機關處理記錄"
         Index           =   10
         Begin VB.Menu mnu0910 
            Caption         =   "來電記錄"
            Index           =   1
         End
         Begin VB.Menu mnu0910 
            Caption         =   "去電記錄"
            Index           =   2
         End
      End
      Begin VB.Menu mnu09 
         Caption         =   "公文來函判發作業"
         Index           =   11
      End
      Begin VB.Menu mnu09 
         Caption         =   "結案單審核作業"
         Index           =   13
         Visible         =   0   'False
      End
      Begin VB.Menu mnu09 
         Caption         =   "P案指示信判發作業"
         Index           =   14
      End
      Begin VB.Menu mnu09 
         Caption         =   "專利處收件夾信件處理"
         Index           =   15
      End
      Begin VB.Menu mnu09 
         Caption         =   "郵件分信關鍵字對照表維護"
         Index           =   16
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
         Begin VB.Menu mnu1101 
            Caption         =   "ＦＭＰ解除期限"
            Index           =   4
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
         Caption         =   "多國案卷號關係建立"
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
               Caption         =   "請款單折扣案件明細"
               Index           =   10
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
      Begin VB.Menu mnu11 
         Caption         =   "變更FC代理人作業"
         Enabled         =   0   'False
         Index           =   20
         Visible         =   0   'False
      End
      Begin VB.Menu mnu11 
         Caption         =   "客戶應收帳款收文檢查上限"
         Index           =   21
         Visible         =   0   'False
      End
      Begin VB.Menu mnu11 
         Caption         =   "客戶預定收款日放寬月數上限"
         Index           =   22
         Visible         =   0   'False
      End
      Begin VB.Menu mnu11 
         Caption         =   "客戶特殊付款週期維護"
         Index           =   23
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "設定"
      Index           =   15
      Begin VB.Menu mnu15 
         Caption         =   "系統印表機設定"
         Index           =   0
      End
      Begin VB.Menu mnu15 
         Caption         =   "報表紙張格式設定"
         Index           =   1
      End
      Begin VB.Menu mnu15 
         Caption         =   "解除畫面擷取限制"
         Index           =   2
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
            Caption         =   "大陸指示信"
            Index           =   7
         End
      End
      Begin VB.Menu mnu16 
         Caption         =   "定稿資料維護"
         Index           =   2
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
         Caption         =   "電話分機資料維護"
         Index           =   5
         Visible         =   0   'False
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
   Begin VB.Menu mnuPopEDoc 
      Caption         =   "電子機關來函彈跳選單"
      Visible         =   0   'False
      Begin VB.Menu mnuPopEDocItem 
         Caption         =   "實審通知日輸入"
         Index           =   1
      End
      Begin VB.Menu mnuPopEDocItem 
         Caption         =   "核准函輸入"
         Index           =   2
      End
      Begin VB.Menu mnuPopEDocItem 
         Caption         =   "核駁函輸入"
         Index           =   3
      End
      Begin VB.Menu mnuPopEDocItem 
         Caption         =   "專利權消滅函輸入"
         Index           =   4
      End
      Begin VB.Menu mnuPopEDocItem 
         Caption         =   "一般來函輸入"
         Index           =   5
      End
      Begin VB.Menu mnuPopEDocItem 
         Caption         =   "證書號數輸入"
         Index           =   6
      End
      Begin VB.Menu mnuPopEDocItem 
         Caption         =   "異議/舉發受理函輸入"
         Index           =   7
      End
   End
   Begin VB.Menu mnuPopEMail1 
      Caption         =   "P案EMail來函彈跳選單"
      Visible         =   0   'False
      Begin VB.Menu mnuPopPEMailItem 
         Caption         =   "申請案號輸入"
         Index           =   1
      End
      Begin VB.Menu mnuPopPEMailItem 
         Caption         =   "審查機關來函"
         Index           =   2
         Begin VB.Menu mnuPopPEMailItem2 
            Caption         =   "實審通知日輸入"
            Index           =   1
         End
         Begin VB.Menu mnuPopPEMailItem2 
            Caption         =   "初審及公佈通知來函輸入"
            Index           =   2
         End
         Begin VB.Menu mnuPopPEMailItem2 
            Caption         =   "核准函輸入"
            Index           =   3
         End
         Begin VB.Menu mnuPopPEMailItem2 
            Caption         =   "核駁函輸入"
            Index           =   4
         End
         Begin VB.Menu mnuPopPEMailItem2 
            Caption         =   "消滅函輸入"
            Index           =   5
         End
         Begin VB.Menu mnuPopPEMailItem2 
            Caption         =   "一般來函輸入"
            Index           =   6
         End
         Begin VB.Menu mnuPopPEMailItem2 
            Caption         =   "證書號數輸入"
            Index           =   7
         End
         Begin VB.Menu mnuPopPEMailItem2 
            Caption         =   "異議/舉發受理函輸入"
            Index           =   8
         End
      End
      Begin VB.Menu mnuPopPEMailItem 
         Caption         =   "代理人來函"
         Index           =   3
         Begin VB.Menu mnuPopPEMailItem3 
            Caption         =   "已收達/已提申"
            Index           =   1
         End
         Begin VB.Menu mnuPopPEMailItem3 
            Caption         =   "通知修正"
            Index           =   2
         End
         Begin VB.Menu mnuPopPEMailItem3 
            Caption         =   "所外鑑定報告結果"
            Index           =   3
         End
         Begin VB.Menu mnuPopPEMailItem3 
            Caption         =   "其他來函"
            Index           =   4
         End
         Begin VB.Menu mnuPopPEMailItem3 
            Caption         =   "已順稿"
            Index           =   5
         End
      End
      Begin VB.Menu mnuPopPEMailItem 
         Caption         =   "其他"
         Index           =   4
         Begin VB.Menu mnuPopPEMailItem4 
            Caption         =   "內部收文"
            Index           =   1
         End
         Begin VB.Menu mnuPopPEMailItem4 
            Caption         =   "年費逾期補繳通知函"
            Index           =   2
         End
         Begin VB.Menu mnuPopPEMailItem4 
            Caption         =   "實審請求期限屆滿前通知函"
            Index           =   3
         End
         Begin VB.Menu mnuPopPEMailItem4 
            Caption         =   "未收文期限提醒E-Mail"
            Index           =   4
         End
         Begin VB.Menu mnuPopPEMailItem4 
            Caption         =   "帳單輸入"
            Index           =   5
         End
         Begin VB.Menu mnuPopPEMailItem4 
            Caption         =   "抵帳單輸入"
            Index           =   6
         End
         Begin VB.Menu mnuPopPEMailItem4 
            Caption         =   "帳單作廢輸入"
            Index           =   7
         End
      End
   End
   Begin VB.Menu mnuPopEMail2 
      Caption         =   "CFP案EMail來函彈跳選單"
      Visible         =   0   'False
      Begin VB.Menu mnuPopCFPEMailItem 
         Caption         =   "代理人案件提申"
         Index           =   1
      End
      Begin VB.Menu mnuPopCFPEMailItem 
         Caption         =   "審查機關來函"
         Index           =   2
         Begin VB.Menu mnuPopCFPEMailItem2 
            Caption         =   "一般來函輸入"
            Index           =   1
         End
         Begin VB.Menu mnuPopCFPEMailItem2 
            Caption         =   "公開/公告資料輸入"
            Index           =   2
         End
         Begin VB.Menu mnuPopCFPEMailItem2 
            Caption         =   "證書號數輸入"
            Index           =   3
         End
         Begin VB.Menu mnuPopCFPEMailItem2 
            Caption         =   "消滅函輸入"
            Index           =   4
         End
         Begin VB.Menu mnuPopCFPEMailItem2 
            Caption         =   "年費逾期補繳通知函"
            Index           =   5
         End
         Begin VB.Menu mnuPopCFPEMailItem2 
            Caption         =   "實體審查、領證費逾期補繳通知函"
            Index           =   6
         End
      End
      Begin VB.Menu mnuPopCFPEMailItem 
         Caption         =   "代理人來函"
         Index           =   3
         Begin VB.Menu mnuPopCFPEMailItem3 
            Caption         =   "已收達"
            Index           =   1
         End
         Begin VB.Menu mnuPopCFPEMailItem3 
            Caption         =   "通知修正"
            Index           =   2
         End
         Begin VB.Menu mnuPopCFPEMailItem3 
            Caption         =   "其他來函"
            Index           =   3
         End
      End
      Begin VB.Menu mnuPopCFPEMailItem 
         Caption         =   "其他"
         Index           =   4
         Begin VB.Menu mnuPopCFPEMailItem4 
            Caption         =   "發文"
            Index           =   1
         End
         Begin VB.Menu mnuPopCFPEMailItem4 
            Caption         =   "內部收文"
            Index           =   2
         End
         Begin VB.Menu mnuPopCFPEMailItem4 
            Caption         =   "期限通知管制表"
            Index           =   3
         End
         Begin VB.Menu mnuPopCFPEMailItem4 
            Caption         =   "未收文期限提醒E-Mail"
            Index           =   4
         End
         Begin VB.Menu mnuPopCFPEMailItem4 
            Caption         =   "帳單輸入"
            Index           =   5
         End
         Begin VB.Menu mnuPopCFPEMailItem4 
            Caption         =   "抵帳單輸入"
            Index           =   6
         End
         Begin VB.Menu mnuPopCFPEMailItem4 
            Caption         =   "帳單作廢輸入"
            Index           =   7
         End
      End
   End
   Begin VB.Menu mnuPopEMail3 
      Caption         =   "FCP案EMail來函彈跳選單"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2005/7/5 整理
Option Explicit

'Add by Morgan 2003/12/23
Dim WithEvents eventConn As ADODB.Connection
Attribute eventConn.VB_VarHelpID = -1
Public bolReOpen As Boolean
'move to basquery
'intPCaseKind分案之系統分類，intPWhere 0國內  1國外CF  2國外FC
'Public intPCaseKind As Integer, intPWhere As Integer
'Add By Cheng 2003/12/19
Dim PLeft1(0 To 7) As Integer
Dim iPrint As Integer
Dim Page As Integer
Dim strTemp1(0 To 7) As String
'End
'Add by Morgan 2008/11/7 是否已經做過
Dim m_blnActivated As Boolean
Public m_wasMaximized As Boolean 'Added by Morgan 畫面最小化後判斷原來是否為最大化用
Dim m_UserNo As String 'Added by Morgan 2015/10/12 報價定稿人員
Dim oControl As Control  'Added by Morgan 2022/1/22
Public Tmpfrm04010519 As Form 'Add By Sindy 2022/5/20
Public m_ChkIsOpenFrm180203 As Boolean 'Add By Sindy 2013/7/8


Function WCmdLog(oStrLog As String)
On Error GoTo ErrHnd

Dim ffa As Integer
ffa = FreeFile
Open App.path & "\cmdlog.log" For Append As ffa
Print #ffa, Trim(Now) & "  ==>  " & oStrLog
Close ffa

ErrHnd:
End Function

'Add by Morgan 2003/12/23
Private Sub eventConn_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
   tmrConnect.Tag = 0
   'Debug.Print Format(Now, "nn:ss:") & Right(Format(Timer, ".00"), 2) & "-->" & pCommand.CommandText
   If strUserNum = "92012" Then WCmdLog pCommand.CommandText 'Added by Morgan 2015/10/26
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

'Add by Morgan 2004/10/14 加控制只要有動作就重算
Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
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

Private Sub mnu040108_Click(Index As Integer)
   
   Select Case Index
   Case 1   '國內外案件資料維護
      If CheckUse("frm050106_1", strExec) = True Then
         frm050106_1.intWhereToGo = 0
         frm050106_1.Show
      End If
   Case 2   '大陸發明案件資料維護
      If CheckUse("frm040108_1", strExec) = True Then
         frm040108_1.Show
      End If
   'Add by Morgan 2004/6/14
   Case 3  '一案兩申請案件資料維護
      If CheckUse("frm040109_1", strExec) = True Then
         frm040109_1.Show
      End If
      
   'Add by Morgan 2015/9/10
   Case 4  '擬制喪失新穎性案件資料維護
      If CheckUse("frm040109_1", strExec) = True Then
         frm040109_1.m_CM10 = "6"
         frm040109_1.Show
      End If
      
   'add by nickc 2005/06/23 大陸香港案
   Case 5  '大陸香港案件資料維護
      If CheckUse("frm050109_1", strExec) = True Then
         frm050109_1.intWhereToGo = 0
         frm050109_1.iK_CM10 = "4" 'Added by Lydia 2015/09/09
         frm050109_1.Show
      End If
      
   'Added by Lydia 2015/09/09
    Case 6 '大陸澳門案
      If CheckUse("frm050109_1", strExec) = True Then
         frm050109_1.intWhereToGo = 0
         frm050109_1.iK_CM10 = "5"
         frm050109_1.Show
      End If
   End Select
End Sub

'Removed by Morgan 2018/4/24
''Modified by Lydia 2015/07/28 大陸澳門案以後,功能表index + 1
''Private Sub mnu040116_Click(Index As Integer)
'Private Sub mnu040117_Click(Index As Integer)
'ToolHide
'intPCaseKind = 專利
'intPWhere = 國內
'Select Case Index
'   'ADD BY SONIA 2014/5/7
'   Case 2 '主管機關來函查詢
'      If CheckUse("frm010008", strExec) = False Then
'         Exit Sub
'      End If
'      frm010008.Show
'   'Added by Morgan 2014/1/14
'   Case 3 '電子公文來函
'      If CheckUse("frm04010516", strExec) = True Then
'         frm04010516.Show
'      End If
'End Select
'End Sub
'
''2014/5/2 add by sonia
''Modified by Lydia 2015/07/28 大陸澳門案以後,功能表index + 1
''Private Sub mnu04011601_Click(Index As Integer)
'Private Sub mnu04011701_Click(Index As Integer)
'   ToolHide
'   If CheckUse("frm010002", strExec) = False Then
'      Exit Sub
'   End If
'   frm010001_1.intChoose = 0
'   frm010001_1.intReceiveKind = 1
'   frm010001_1.intModifyKind = Index - 1
'   Select Case Index
'      Case 1
'         If CheckUse("frm010002", strAdd) = False Then
'            Exit Sub
'         End If
'         '新增：直接跳至Ckind
'         frm010002.Caption = "主管機關來函－新增"
'         frm010002.lblRecieveCode.Caption = "D" + CompAutoNumberYear(GetTaiwanThisYear)
'      Case 2
'         If CheckUse("frm010002", strEdit) = False Then
'            Exit Sub
'         End If
'         frm010001_1.Caption = "主管機關來函－修改"
'      Case 3
'         frm010001_1.Caption = "主管機關來函－查詢"
'   End Select
'
'End Sub
''2014/5/2 end
'end 2018/4/24

'Add by Morgan 2005/8/25 通知函
Private Sub mnu040301_Click(Index As Integer)
   intPCaseKind = 專利
   intPWhere = 國內
   ToolHide
   Select Case Index
      Case 1   '公開通知函
         If CheckUse("frm040325", strExec) = True Then
            frm040325.Show
         End If
      Case 2   '公告通知函
         If CheckUse("frm040301", strExec) = True Then
            frm040301.Show
         End If
      Case 3   '公告期滿通知函
         If CheckUse("frm040302", strExec) = True Then
            frm040302.Show
         End If
      Case 4   '繳年費/實體審查通知函
         If CheckUse("frm040303", strExec) = True Then
            frm040303.Show
         End If
         
      
      '92.3.5 Add By sonia
      'Removed by Morgan 2020/1/16 移到機關來函
      'Case 5   '年費逾期補繳通知函
      '   If CheckUse("frm040324", strExec) = True Then
      '      frm040324.iKind = 1 'Added by Lydia 2015/07/20
      '      frm040324.Show
      '   End If
      'end 2020/1/16
      '92.3.5 end
      
      'Add By Cheng 2002/06/24
      Case 6   '其他通知函/聯絡單
         If CheckUse("frm040322", strExec) = True Then
            frm040322.Show
         End If
      Case 7   '年費逾期補繳通知函(整批)
         If CheckUse("frm040326", strExec) = True Then
            frm040326.Show
         End If
         
      'Added by Lydia 2015/07/20
      'Removed by Morgan 2020/1/16 移到機關來函
      'Case 8   '實審請求期限屆滿前通知函-共用表單
      '   If CheckUse("frm040324", strExec) = True Then
      '      frm040324.iKind = 2
      '      frm040324.Show
      '   End If
      'end 2020/1/16
      
      'Added by Morgan 2015/9/2
      Case 9 '期限通知檢核及報表列印
         If CheckUse("frm040335", strExec) = True Then
            frm040335.Show
         End If
      'add by sonia 2016/1/11
      Case 10  '未收文期限提醒E-Mail
         If CheckUse("frm050326", strExec) = True Then
            frm050326.Show
            frm050326.Text1(1) = "P"
         End If
   End Select
End Sub

Private Sub mnu0406_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 4   '公開及公告市場統計表
         If CheckUse("frm04060108", strExec) = True Then
            frm04060108.Show
         End If
      'Add By Sindy 2013/8/29
      Case 5   '公報產業分類案件市佔分析
         If CheckUse("frm100133", strExec) = True Then
            frm100133.Show
         End If
     Case Else
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
      'Added by Lydia 2023/01/03
      Case 4 '專業部主管分案作業
         If CheckUse("frm210156", strExec) = True Then
            frm210156.Show
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
      'add by Sindy 2014/4/18
      Case 5 '電話分機資料維護
         'Mark by Amy 2023/08/17 表單數已達最大限制
'         If CheckUse("frm010028", strExec) = True Then
'            frm010028.Show
'         End If
      'Add By Sindy 2016/3/21
      Case 7 '系統收件區
         'Modify By Sindy 2022/5/13
         If Left(Pub_StrUserSt03, 2) = "F2" Then  '外專
            frm06010616.Show
         Else
         '2022/5/13 END
            'Modify By Sindy 2017/12/20
'            If Left(Pub_StrUserSt03, 2) = "P1" Then '內專
               frm04010519.Show
'            End If
'            '2017/12/20 END
         End If
         
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
'         Call frm160201.cmdok_Click(0)
         frm180203_1.Show
      Case 3 '員工個人資料明細確認
         frm160102.intChoose = 1
         frm160102.Hide
         Call frm160102.cmdOK_Click(0)
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

'Add by Amy 2014/05/22
Private Sub mnuChUser_Click()
   frmChgUser.Show
End Sub

Private Sub mnuDML_Click(Index As Integer)
    frmDML.Show
End Sub

'Added by Morgan 2014/1/14
Private Sub mnuPopEDocItem_Click(Index As Integer)
   frm04010516.OpenForm Index
End Sub

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

'Added by Sindy 2016/9/20
Private Sub mnuPopPEMailItem_Click(Index As Integer)
   OpenForm_P1 Index
End Sub
Private Sub mnuPopPEMailItem2_Click(Index As Integer)
   OpenForm_P2 Index
End Sub
Private Sub mnuPopPEMailItem3_Click(Index As Integer)
   OpenForm_P3 Index
End Sub
Private Sub mnuPopPEMailItem4_Click(Index As Integer)
   OpenForm_P4 Index
End Sub
Private Sub mnuPopCFPEMailItem_Click(Index As Integer)
   OpenForm_CFP1 Index
End Sub
Private Sub mnuPopCFPEMailItem2_Click(Index As Integer)
   OpenForm_CFP2 Index
End Sub
Private Sub mnuPopCFPEMailItem3_Click(Index As Integer)
   OpenForm_CFP3 Index
End Sub
Private Sub mnuPopCFPEMailItem4_Click(Index As Integer)
   OpenForm_CFP4 Index
End Sub
'2016/9/20 END

'Add By Sindy 2022/5/20
Public Sub SetTmpfrm04010519(tmpForm As Form)
   Set Tmpfrm04010519 = tmpForm
End Sub

'Add By Sindy 2019/4/24
Private Sub OpenForm_P1(Index As Integer)
   Select Case Index
      Case 1 '申請案號輸入
         If CheckUse("frm04010401", strExec) = True Then
            frm04010401.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm04010401.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm04010401.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm04010401.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm04010401.m_strCP01 = Tmpfrm04010519.txtPI18
            frm04010401.m_strCP02 = Tmpfrm04010519.txtPI19
            frm04010401.m_strCP03 = Tmpfrm04010519.txtPI20
            frm04010401.m_strCP04 = Tmpfrm04010519.txtPI21
            frm04010401.m_RDate = Tmpfrm04010519.m_strPi12
            frm04010401.Show
         End If
   End Select
End Sub
Private Sub OpenForm_P2(Index As Integer)
Dim strDocNo As String, strRecDate As String, strDocWord As String, strDeadLine As String
   
'   If m_strPA11 = "" Then
'      MsgBox "無申請案號，不可執行此作業！", vbExclamation, "警告！"
'      Exit Sub
'   End If
   
   intPCaseKind = 專利
   intPWhere = 國內
   Select Case Index
      Case 1 '實審通知日輸入
         If CheckUse("frm04010501", strExec) = True Then
            frm04010501.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm04010501.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm04010501.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm04010501.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm04010501.m_strCP01 = Tmpfrm04010519.txtPI18
            frm04010501.m_strCP02 = Tmpfrm04010519.txtPI19
            frm04010501.m_strCP03 = Tmpfrm04010519.txtPI20
            frm04010501.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm04010501.m_AppNo = m_strPA11
            frm04010501.m_RDate = Tmpfrm04010519.m_strPi12
            frm04010501.Show
         End If
      Case 2 '初審及公佈通知來函輸入
         If CheckUse("frm04010514_1", strExec) = True Then
            frm04010514_1.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm04010514_1.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm04010514_1.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm04010514_1.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm04010514_1.m_strCP01 = Tmpfrm04010519.txtPI18
            frm04010514_1.m_strCP02 = Tmpfrm04010519.txtPI19
            frm04010514_1.m_strCP03 = Tmpfrm04010519.txtPI20
            frm04010514_1.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm04010514_1.m_AppNo = m_strPA11
            frm04010514_1.m_RDate = Tmpfrm04010519.m_strPi12
            frm04010514_1.Show
         End If
      Case 3 '核准函輸入
         If CheckUse("frm04010502_1", strExec) = True Then
            frm04010502_1.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm04010502_1.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm04010502_1.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm04010502_1.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm04010502_1.m_strCP01 = Tmpfrm04010519.txtPI18
            frm04010502_1.m_strCP02 = Tmpfrm04010519.txtPI19
            frm04010502_1.m_strCP03 = Tmpfrm04010519.txtPI20
            frm04010502_1.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm04010502_1.m_AppNo = m_strPA11
            frm04010502_1.m_RDate = Tmpfrm04010519.m_strPi12
            'frm04010502_1.m_DeadLine = strDeadLine
            frm04010502_1.Show
         End If
      Case 4 '核駁函輸入
         If CheckUse("frm04010503_1", strExec) = True Then
            frm04010503_1.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm04010503_1.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm04010503_1.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm04010503_1.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm04010503_1.m_strCP01 = Tmpfrm04010519.txtPI18
            frm04010503_1.m_strCP02 = Tmpfrm04010519.txtPI19
            frm04010503_1.m_strCP03 = Tmpfrm04010519.txtPI20
            frm04010503_1.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm04010503_1.m_AppNo = m_strPA11
            frm04010503_1.m_RDate = Tmpfrm04010519.m_strPi12
            'frm04010503_1.m_DeadLine = strDeadLine
            frm04010503_1.Show
         End If
      Case 5 '消滅函輸入
         If CheckUse("frm04010511_1", strExec) = True Then
            frm04010511_1.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm04010511_1.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm04010511_1.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm04010511_1.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm04010511_1.m_strCP01 = Tmpfrm04010519.txtPI18
            frm04010511_1.m_strCP02 = Tmpfrm04010519.txtPI19
            frm04010511_1.m_strCP03 = Tmpfrm04010519.txtPI20
            frm04010511_1.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm04010511_1.m_AppNo = m_strPA11
            frm04010511_1.m_RDate = Tmpfrm04010519.m_strPi12
            frm04010511_1.Show
         End If
      Case 6 '一般來函輸入
         If CheckUse("frm04010504_1", strExec) = True Then
            frm04010504_1.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm04010504_1.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm04010504_1.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm04010504_1.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm04010504_1.m_strCP01 = Tmpfrm04010519.txtPI18
            frm04010504_1.m_strCP02 = Tmpfrm04010519.txtPI19
            frm04010504_1.m_strCP03 = Tmpfrm04010519.txtPI20
            frm04010504_1.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm04010504_1.m_AppNo = m_strPA11
            frm04010504_1.m_RDate = Tmpfrm04010519.m_strPi12
            'frm04010504_1.m_DeadLine = strDeadLine
            ''frm04010504_1.m_NewCP10 = pCP10
            frm04010504_1.Show
         End If
      Case 7 '證書號數輸入
         If CheckUse("frm04010505_1", strExec) = True Then
            frm04010505_1.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm04010505_1.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm04010505_1.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm04010505_1.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm04010505_1.m_strCP01 = Tmpfrm04010519.txtPI18
            frm04010505_1.m_strCP02 = Tmpfrm04010519.txtPI19
            frm04010505_1.m_strCP03 = Tmpfrm04010519.txtPI20
            frm04010505_1.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm04010505_1.m_AppNo = m_strPA11
            frm04010505_1.m_RDate = Tmpfrm04010519.m_strPi12
            frm04010505_1.Show
         End If
      Case 8 '異議/舉發受理函輸入
         If CheckUse("frm04010506_1", strExec) = True Then
            frm04010506_1.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm04010506_1.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm04010506_1.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm04010506_1.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm04010506_1.m_strCP01 = Tmpfrm04010519.txtPI18
            frm04010506_1.m_strCP02 = Tmpfrm04010519.txtPI19
            frm04010506_1.m_strCP03 = Tmpfrm04010519.txtPI20
            frm04010506_1.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm04010506_1.m_AppNo = m_strPA11
            frm04010506_1.m_RDate = Tmpfrm04010519.m_strPi12
            frm04010506_1.Show
         End If
   End Select
End Sub
Private Sub OpenForm_P3(Index As Integer)
   intPCaseKind = 專利
   intPWhere = 國內
   Select Case Index
      Case 1 '代理人已收達/己提申
         If CheckUse("frm04010507_1", strExec) = True Then
            frm04010507_1.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm04010507_1.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm04010507_1.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm04010507_1.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm04010507_1.m_strCP01 = Tmpfrm04010519.txtPI18
            frm04010507_1.m_strCP02 = Tmpfrm04010519.txtPI19
            frm04010507_1.m_strCP03 = Tmpfrm04010519.txtPI20
            frm04010507_1.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm04010507_1.m_AppNo = m_strPA11
            frm04010507_1.m_RDate = Tmpfrm04010519.m_strPi12
            frm04010507_1.Show
         End If
      Case 2 '代理人通知修正
         If CheckUse("frm04010508_1", strExec) = True Then
            frm04010508_1.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm04010508_1.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm04010508_1.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm04010508_1.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm04010508_1.m_strCP01 = Tmpfrm04010519.txtPI18
            frm04010508_1.m_strCP02 = Tmpfrm04010519.txtPI19
            frm04010508_1.m_strCP03 = Tmpfrm04010519.txtPI20
            frm04010508_1.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm04010508_1.m_AppNo = m_strPA11
            frm04010508_1.m_RDate = Tmpfrm04010519.m_strPi12
            frm04010508_1.Show
         End If
      Case 3 '所外鑑定報告結果
         If CheckUse("frm04010509_1", strExec) = True Then
            frm04010509_1.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm04010509_1.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm04010509_1.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm04010509_1.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm04010509_1.m_strCP01 = Tmpfrm04010519.txtPI18
            frm04010509_1.m_strCP02 = Tmpfrm04010519.txtPI19
            frm04010509_1.m_strCP03 = Tmpfrm04010519.txtPI20
            frm04010509_1.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm04010509_1.m_AppNo = m_strPA11
            frm04010509_1.m_RDate = Tmpfrm04010519.m_strPi12
            frm04010509_1.Show
         End If
      Case 4 '其他來函輸入
         If CheckUse("frm02010603_1", strExec) = True Then
            Call frm02010603_1.SetParent(Tmpfrm04010519)
            frm02010603_1.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm02010603_1.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm02010603_1.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm02010603_1.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm02010603_1.m_strCP01 = Tmpfrm04010519.txtPI18
            frm02010603_1.m_strCP02 = Tmpfrm04010519.txtPI19
            frm02010603_1.m_strCP03 = Tmpfrm04010519.txtPI20
            frm02010603_1.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm02010603_1.m_AppNo = m_strPA11
            frm02010603_1.m_RDate = Tmpfrm04010519.m_strPi12
            frm02010603_1.Caption = "其他來函輸入"
            frm02010603_1.Show
         End If
      Case 5 '已順稿
         If CheckUse("frm04010513", strExec) = True Then
            frm04010513.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm04010513.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm04010513.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm04010513.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm04010513.m_strCP01 = Tmpfrm04010519.txtPI18
            frm04010513.m_strCP02 = Tmpfrm04010519.txtPI19
            frm04010513.m_strCP03 = Tmpfrm04010519.txtPI20
            frm04010513.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm04010513.m_AppNo = m_strPA11
            frm04010513.m_RDate = Tmpfrm04010519.m_strPi12
            frm04010513.Show
         End If
   End Select
End Sub
Private Sub OpenForm_P4(Index As Integer)
   intPCaseKind = 專利
   intPWhere = 國內
   Select Case Index
      Case 1 '內部收文
         If CheckUse("frm010001", strExec) = True Then
            Call frm010001.SetParent(Tmpfrm04010519) 'Modify By Sindy 2020/5/27
            frm010001.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm010001.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm010001.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm010001.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm010001.m_strCP01 = Tmpfrm04010519.txtPI18
            frm010001.m_strCP02 = Tmpfrm04010519.txtPI19
            frm010001.m_strCP03 = Tmpfrm04010519.txtPI20
            frm010001.m_strCP04 = Tmpfrm04010519.txtPI21
            frm010001.m_RDate = Tmpfrm04010519.m_strPi12
            'Set frm010001.mPrevForm = Tmpfrm04010519
            frm010001.intChoose = 1
            frm010001.intReceiveKind = 0
            frm010001.intModifyKind = 0
            frm010001.Caption = "內部收文－新增"
         End If
      Case 2 '年費逾期補繳通知函
         If CheckUse("frm040324", strExec) = True Then
            frm040324.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm040324.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm040324.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm040324.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm040324.m_strCP01 = Tmpfrm04010519.txtPI18
            frm040324.m_strCP02 = Tmpfrm04010519.txtPI19
            frm040324.m_strCP03 = Tmpfrm04010519.txtPI20
            frm040324.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm040324.m_AppNo = m_strPA11
            frm040324.m_RDate = Tmpfrm04010519.m_strPi12
            frm040324.iKind = 1
            frm040324.Show
         End If
      Case 3 '實審請求期限屆滿前通知函
         If CheckUse("frm040324", strExec) = True Then
            frm040324.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm040324.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm040324.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm040324.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm040324.m_strCP01 = Tmpfrm04010519.txtPI18
            frm040324.m_strCP02 = Tmpfrm04010519.txtPI19
            frm040324.m_strCP03 = Tmpfrm04010519.txtPI20
            frm040324.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm040324.m_AppNo = m_strPA11
            frm040324.m_RDate = Tmpfrm04010519.m_strPi12
            frm040324.iKind = 2
            frm040324.Show
         End If
      Case 4 '未收文期限提醒E-Mail
         If CheckUse("frm050326", strExec) = True Then
            frm050326.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm050326.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm050326.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm050326.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm050326.m_strCP01 = Tmpfrm04010519.txtPI18
            frm050326.m_strCP02 = Tmpfrm04010519.txtPI19
            frm050326.m_strCP03 = Tmpfrm04010519.txtPI20
            frm050326.m_strCP04 = Tmpfrm04010519.txtPI21
            Set frm050326.m_PrevForm = Tmpfrm04010519
            frm050326.Show
            frm050326.Text1(1) = "P"
         End If
      'Add By Sindy 2018/2/23
      Case 5 '帳單輸入
         'Add By Sindy 2025/4/15 給外專人員操作寰華案
         'Removed by Moran 2025/8/14 配合寰華案帳單電子化調整--Sharon
         'If Left(Pub_StrUserSt15, 2) = "F2" Then
         '   If CheckUse("Frmacc2150", strExec) = True Then
         '      ToolShow
         '      tool1_enabled
         '      Frmacc2150.m_strIR01 = Tmpfrm04010519.m_strIR01
         '      Frmacc2150.m_strIR02 = Tmpfrm04010519.m_strIR02
         '      Frmacc2150.m_strIR03 = Tmpfrm04010519.m_strIR03
         '      Frmacc2150.m_strIR04 = Tmpfrm04010519.m_strIR04
         '      Frmacc2150.m_CP01 = Tmpfrm04010519.txtPI18
         '      Frmacc2150.m_CP02 = Tmpfrm04010519.txtPI19
         '      Frmacc2150.m_CP03 = Tmpfrm04010519.txtPI20
         '      Frmacc2150.m_CP04 = Tmpfrm04010519.txtPI21
         '      Frmacc2150.m_RDate = Tmpfrm04010519.m_strPi12
         '      Set Frmacc2150.m_ParentForm = Tmpfrm04010519
         '      Frmacc2150.Show
         '   End If
         'Else
         'end 2025/8/14
         '2025/4/15 END
         
            If CheckUse("Frmacc21u0", strExec) = True Then
               ToolShow
               'tool1_enabled
               tool3_enabled
               Frmacc21u0.m_strIR01 = Tmpfrm04010519.m_strIR01
               Frmacc21u0.m_strIR02 = Tmpfrm04010519.m_strIR02
               Frmacc21u0.m_strIR03 = Tmpfrm04010519.m_strIR03
               Frmacc21u0.m_strIR04 = Tmpfrm04010519.m_strIR04
               Frmacc21u0.m_strCP01 = Tmpfrm04010519.txtPI18
               Frmacc21u0.m_strCP02 = Tmpfrm04010519.txtPI19
               Frmacc21u0.m_strCP03 = Tmpfrm04010519.txtPI20
               Frmacc21u0.m_strCP04 = Tmpfrm04010519.txtPI21
               Frmacc21u0.m_RDate = Tmpfrm04010519.m_strPi12
               Set Frmacc21u0.m_PrevForm = Tmpfrm04010519
               Frmacc21u0.Show
            End If
            
         'End If 'Removed by Morgan 2025/8/14
      'Add By Sindy 2018/2/23
      Case 6 '抵帳單輸入
         If CheckUse("Frmacc2160", strExec) = True Then
            ToolShow
            tool1_enabled
            Frmacc2160.m_strIR01 = Tmpfrm04010519.m_strIR01
            Frmacc2160.m_strIR02 = Tmpfrm04010519.m_strIR02
            Frmacc2160.m_strIR03 = Tmpfrm04010519.m_strIR03
            Frmacc2160.m_strIR04 = Tmpfrm04010519.m_strIR04
            Frmacc2160.m_strCP01 = Tmpfrm04010519.txtPI18
            Frmacc2160.m_strCP02 = Tmpfrm04010519.txtPI19
            Frmacc2160.m_strCP03 = Tmpfrm04010519.txtPI20
            Frmacc2160.m_strCP04 = Tmpfrm04010519.txtPI21
            Frmacc2160.m_RDate = Tmpfrm04010519.m_strPi12
            Set Frmacc2160.m_PrevForm = Tmpfrm04010519
            Frmacc2160.Show
         End If
      'Add By Sindy 2018/2/23
      Case 7 '帳單作廢輸入
         If CheckUse("Frmacc21j0", strExec) = True Then
            ToolShow
            tool1_enabled
            Frmacc21j0.m_strIR01 = Tmpfrm04010519.m_strIR01
            Frmacc21j0.m_strIR02 = Tmpfrm04010519.m_strIR02
            Frmacc21j0.m_strIR03 = Tmpfrm04010519.m_strIR03
            Frmacc21j0.m_strIR04 = Tmpfrm04010519.m_strIR04
            Frmacc21j0.m_strCP01 = Tmpfrm04010519.txtPI18
            Frmacc21j0.m_strCP02 = Tmpfrm04010519.txtPI19
            Frmacc21j0.m_strCP03 = Tmpfrm04010519.txtPI20
            Frmacc21j0.m_strCP04 = Tmpfrm04010519.txtPI21
            Frmacc21j0.m_RDate = Tmpfrm04010519.m_strPi12
            Set Frmacc21j0.m_PrevForm = Tmpfrm04010519
            Frmacc21j0.Show
         End If
   End Select
End Sub
Private Sub OpenForm_CFP1(Index As Integer)
   Select Case Index
      Case 1 '代理人案件提申
         If CheckUse("frm050103_1", strExec) = True Then
            frm050103_1.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm050103_1.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm050103_1.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm050103_1.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm050103_1.m_strCP01 = Tmpfrm04010519.txtPI18
            frm050103_1.m_strCP02 = Tmpfrm04010519.txtPI19
            frm050103_1.m_strCP03 = Tmpfrm04010519.txtPI20
            frm050103_1.m_strCP04 = Tmpfrm04010519.txtPI21
            frm050103_1.m_RDate = Tmpfrm04010519.m_strPi12
            frm050103_1.Show
         End If
   End Select
End Sub
Private Sub OpenForm_CFP2(Index As Integer)
   intPCaseKind = 專利
   intPWhere = 國外_CF
   Select Case Index
      Case 1 '一般來函
         If CheckUse("frm05010401_1", strExec) = True Then
            frm05010401_1.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm05010401_1.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm05010401_1.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm05010401_1.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm05010401_1.m_strCP01 = Tmpfrm04010519.txtPI18
            frm05010401_1.m_strCP02 = Tmpfrm04010519.txtPI19
            frm05010401_1.m_strCP03 = Tmpfrm04010519.txtPI20
            frm05010401_1.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm05010401_1.m_AppNo = m_strPA11
            frm05010401_1.m_RDate = Tmpfrm04010519.m_strPi12
            frm05010401_1.Caption = "一般來函輸入"
            frm05010401_1.Show
         End If
      Case 2 '公開公告資料輸入
         If CheckUse("frm05010402_1", strExec) = True Then
            frm05010402_1.intChoose = 1
            frm05010402_1.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm05010402_1.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm05010402_1.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm05010402_1.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm05010402_1.m_strCP01 = Tmpfrm04010519.txtPI18
            frm05010402_1.m_strCP02 = Tmpfrm04010519.txtPI19
            frm05010402_1.m_strCP03 = Tmpfrm04010519.txtPI20
            frm05010402_1.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm05010402_1.m_AppNo = m_strPA11
            frm05010402_1.m_RDate = Tmpfrm04010519.m_strPi12
            frm05010402_1.Caption = "公開公告資料輸入"
            frm05010402_1.Show
         End If
      Case 3 '證書號數輸入
         If CheckUse("frm05010402_1", strExec) = True Then
            frm05010402_1.intChoose = 2
            frm05010402_1.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm05010402_1.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm05010402_1.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm05010402_1.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm05010402_1.m_strCP01 = Tmpfrm04010519.txtPI18
            frm05010402_1.m_strCP02 = Tmpfrm04010519.txtPI19
            frm05010402_1.m_strCP03 = Tmpfrm04010519.txtPI20
            frm05010402_1.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm05010402_1.m_AppNo = m_strPA11
            frm05010402_1.m_RDate = Tmpfrm04010519.m_strPi12
            frm05010402_1.Caption = "證書號數輸入"
            frm05010402_1.Show
         End If
      Case 4 '消滅函輸入
         If CheckUse("frm05010404_1", strExec) = True Then
            frm05010404_1.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm05010404_1.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm05010404_1.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm05010404_1.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm05010404_1.m_strCP01 = Tmpfrm04010519.txtPI18
            frm05010404_1.m_strCP02 = Tmpfrm04010519.txtPI19
            frm05010404_1.m_strCP03 = Tmpfrm04010519.txtPI20
            frm05010404_1.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm05010404_1.m_AppNo = m_strPA11
            frm05010404_1.m_RDate = Tmpfrm04010519.m_strPi12
            frm05010404_1.Caption = "消滅函輸入"
            frm05010404_1.Show
         End If
      Case 5 '年費逾期補繳通知函
         If CheckUse("frm05010405_1", strExec) = True Then
            frm05010405_1.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm05010405_1.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm05010405_1.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm05010405_1.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm05010405_1.m_strCP01 = Tmpfrm04010519.txtPI18
            frm05010405_1.m_strCP02 = Tmpfrm04010519.txtPI19
            frm05010405_1.m_strCP03 = Tmpfrm04010519.txtPI20
            frm05010405_1.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm05010405_1.m_AppNo = m_strPA11
            frm05010405_1.m_RDate = Tmpfrm04010519.m_strPi12
            frm05010405_1.Show
         End If
      Case 6 '實體審查、領證費逾期補繳通知函
         If CheckUse("frm05010406_1", strExec) = True Then
            frm05010406_1.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm05010406_1.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm05010406_1.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm05010406_1.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm05010406_1.m_strCP01 = Tmpfrm04010519.txtPI18
            frm05010406_1.m_strCP02 = Tmpfrm04010519.txtPI19
            frm05010406_1.m_strCP03 = Tmpfrm04010519.txtPI20
            frm05010406_1.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm05010406_1.m_AppNo = m_strPA11
            frm05010406_1.m_RDate = Tmpfrm04010519.m_strPi12
            frm05010406_1.Show
         End If
   End Select
End Sub
Private Sub OpenForm_CFP3(Index As Integer)
   intPCaseKind = 專利
   intPWhere = 國外_CF
   Select Case Index
      Case 1 '代理人已收達/已提申
         If CheckUse("frm02010601_1", strExec) = True Then
            Call frm02010601_1.SetParent(Tmpfrm04010519)
            frm02010601_1.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm02010601_1.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm02010601_1.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm02010601_1.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm02010601_1.m_strCP01 = Tmpfrm04010519.txtPI18
            frm02010601_1.m_strCP02 = Tmpfrm04010519.txtPI19
            frm02010601_1.m_strCP03 = Tmpfrm04010519.txtPI20
            frm02010601_1.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm02010601_1.m_AppNo = m_strPA11
            frm02010601_1.m_RDate = Tmpfrm04010519.m_strPi12
            frm02010601_1.Show
         End If
      Case 2 '代理人通知修正
         If CheckUse("frm02010602_1", strExec) = True Then
            Call frm02010602_1.SetParent(Tmpfrm04010519)
            frm02010602_1.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm02010602_1.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm02010602_1.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm02010602_1.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm02010602_1.m_strCP01 = Tmpfrm04010519.txtPI18
            frm02010602_1.m_strCP02 = Tmpfrm04010519.txtPI19
            frm02010602_1.m_strCP03 = Tmpfrm04010519.txtPI20
            frm02010602_1.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm02010602_1.m_AppNo = m_strPA11
            frm02010602_1.m_RDate = Tmpfrm04010519.m_strPi12
            frm02010602_1.Show
         End If
      Case 3 '代理人其他來函輸入
         If CheckUse("frm02010603_1", strExec) = True Then
            Call frm02010603_1.SetParent(Tmpfrm04010519)
            frm02010603_1.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm02010603_1.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm02010603_1.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm02010603_1.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm02010603_1.m_strCP01 = Tmpfrm04010519.txtPI18
            frm02010603_1.m_strCP02 = Tmpfrm04010519.txtPI19
            frm02010603_1.m_strCP03 = Tmpfrm04010519.txtPI20
            frm02010603_1.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm02010603_1.m_AppNo = m_strPA11
            frm02010603_1.m_RDate = Tmpfrm04010519.m_strPi12
            frm02010603_1.Caption = "其他來函輸入"
            frm02010603_1.Show
         End If
   End Select
End Sub
Private Sub OpenForm_CFP4(Index As Integer)
   intPCaseKind = 專利
   intPWhere = 國外_CF
   Select Case Index
      Case 1 '發文
         If CheckUse("frm050102_1", strExec) = True Then
            'Call frm050102_1.SetParent(Tmpfrm04010519)
            frm050102_1.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm050102_1.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm050102_1.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm050102_1.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm050102_1.m_strCP01 = Tmpfrm04010519.txtPI18
            frm050102_1.m_strCP02 = Tmpfrm04010519.txtPI19
            frm050102_1.m_strCP03 = Tmpfrm04010519.txtPI20
            frm050102_1.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm050102_1.m_AppNo = m_strPA11
            frm050102_1.m_RDate = Tmpfrm04010519.m_strPi12
            frm050102_1.Show
         End If
      Case 2 '內部收文
         If CheckUse("frm010001", strExec) = True Then
            frm010001.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm010001.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm010001.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm010001.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm010001.m_strCP01 = Tmpfrm04010519.txtPI18
            frm010001.m_strCP02 = Tmpfrm04010519.txtPI19
            frm010001.m_strCP03 = Tmpfrm04010519.txtPI20
            frm010001.m_strCP04 = Tmpfrm04010519.txtPI21
            frm010001.m_RDate = Tmpfrm04010519.m_strPi12
            Set frm010001.mPrevForm = Tmpfrm04010519
            frm010001.intChoose = 1
            frm010001.intReceiveKind = 0
            frm010001.intModifyKind = 0
            frm010001.Caption = "內部收文－新增"
         End If
      Case 3 '期限通知管制表
         If CheckUse("frm050312", strExec) = True Then
            frm050312.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm050312.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm050312.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm050312.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm050312.m_strCP01 = Tmpfrm04010519.txtPI18
            frm050312.m_strCP02 = Tmpfrm04010519.txtPI19
            frm050312.m_strCP03 = Tmpfrm04010519.txtPI20
            frm050312.m_strCP04 = Tmpfrm04010519.txtPI21
'            frm050312.m_AppNo = m_strPA11
            frm050312.m_RDate = Tmpfrm04010519.m_strPi12
            frm050312.Show
         End If
      'Add By Sindy 2018/2/23
      Case 4 '未收文期限提醒E-Mail
         If CheckUse("frm050326", strExec) = True Then
            frm050326.m_strIR01 = Tmpfrm04010519.m_strIR01
            frm050326.m_strIR02 = Tmpfrm04010519.m_strIR02
            frm050326.m_strIR03 = Tmpfrm04010519.m_strIR03
            frm050326.m_strIR04 = Tmpfrm04010519.m_strIR04
            frm050326.m_strCP01 = Tmpfrm04010519.txtPI18
            frm050326.m_strCP02 = Tmpfrm04010519.txtPI19
            frm050326.m_strCP03 = Tmpfrm04010519.txtPI20
            frm050326.m_strCP04 = Tmpfrm04010519.txtPI21
            Set frm050326.m_PrevForm = Tmpfrm04010519
            frm050326.Show
            frm050326.Text1(1) = "CFP"
         End If
'      'Add By Sindy 2018/2/23
'      Case 5 '帳單輸入
'         If CheckUse("Frmacc2150", strExec) = True Then
'            ToolShow
'            tool1_enabled
'            Frmacc2150.m_strIR01 = m_strIR01
'            Frmacc2150.m_strIR02 = m_strIR02
'            Frmacc2150.m_strIR03 = m_strIR03
'            Frmacc2150.m_strIR04 = m_strIR04
'            Frmacc2150.m_CP01 = txtPI18
'            Frmacc2150.m_CP02 = txtPI19
'            Frmacc2150.m_CP03 = txtPI20
'            Frmacc2150.m_CP04 = txtPI21
'            Frmacc2150.m_RDate = m_strPi12
'            Set Frmacc2150.m_ParentForm = Me
'            Frmacc2150.Show
'         End If
      'Add By Sindy 2018/10/19
      Case 5 '帳單輸入
         If CheckUse("Frmacc21u0", strExec) = True Then
            ToolShow
            'tool1_enabled
            tool3_enabled
            Frmacc21u0.m_strIR01 = Tmpfrm04010519.m_strIR01
            Frmacc21u0.m_strIR02 = Tmpfrm04010519.m_strIR02
            Frmacc21u0.m_strIR03 = Tmpfrm04010519.m_strIR03
            Frmacc21u0.m_strIR04 = Tmpfrm04010519.m_strIR04
            Frmacc21u0.m_strCP01 = Tmpfrm04010519.txtPI18
            Frmacc21u0.m_strCP02 = Tmpfrm04010519.txtPI19
            Frmacc21u0.m_strCP03 = Tmpfrm04010519.txtPI20
            Frmacc21u0.m_strCP04 = Tmpfrm04010519.txtPI21
            Frmacc21u0.m_RDate = Tmpfrm04010519.m_strPi12
            Set Frmacc21u0.m_PrevForm = Tmpfrm04010519
            Frmacc21u0.Show
         End If
      'Add By Sindy 2018/2/23
      Case 6 '抵帳單輸入
         If CheckUse("Frmacc2160", strExec) = True Then
            ToolShow
            tool1_enabled
            Frmacc2160.m_strIR01 = Tmpfrm04010519.m_strIR01
            Frmacc2160.m_strIR02 = Tmpfrm04010519.m_strIR02
            Frmacc2160.m_strIR03 = Tmpfrm04010519.m_strIR03
            Frmacc2160.m_strIR04 = Tmpfrm04010519.m_strIR04
            Frmacc2160.m_strCP01 = Tmpfrm04010519.txtPI18
            Frmacc2160.m_strCP02 = Tmpfrm04010519.txtPI19
            Frmacc2160.m_strCP03 = Tmpfrm04010519.txtPI20
            Frmacc2160.m_strCP04 = Tmpfrm04010519.txtPI21
            Frmacc2160.m_RDate = Tmpfrm04010519.m_strPi12
            Set Frmacc2160.m_PrevForm = Tmpfrm04010519
            Frmacc2160.Show
         End If
      'Add By Sindy 2018/2/23
      Case 7 '帳單作廢輸入
         If CheckUse("Frmacc21j0", strExec) = True Then
            ToolShow
            tool1_enabled
            Frmacc21j0.m_strIR01 = Tmpfrm04010519.m_strIR01
            Frmacc21j0.m_strIR02 = Tmpfrm04010519.m_strIR02
            Frmacc21j0.m_strIR03 = Tmpfrm04010519.m_strIR03
            Frmacc21j0.m_strIR04 = Tmpfrm04010519.m_strIR04
            Frmacc21j0.m_strCP01 = Tmpfrm04010519.txtPI18
            Frmacc21j0.m_strCP02 = Tmpfrm04010519.txtPI19
            Frmacc21j0.m_strCP03 = Tmpfrm04010519.txtPI20
            Frmacc21j0.m_strCP04 = Tmpfrm04010519.txtPI21
            Frmacc21j0.m_RDate = Tmpfrm04010519.m_strPi12
            Set Frmacc21j0.m_PrevForm = Tmpfrm04010519
            Frmacc21j0.Show
         End If
   End Select
End Sub
'2019/4/24 END

'Add by Morgan 2005/3/2 控制不可拷貝畫面
Private Sub Timer3_Timer()

   Static dtNow As Date 'Added by Morgan 2024/8/7
      
On Error Resume Next 'Added by Morgan 2017/8/29 若有其他軟體也在使用剪貼簿時會發生521(無法開啟剪貼簿)的錯誤(Ex.Word開啟剪貼簿並擷取畫面)

   'Added by Morgan 2024/8/7 定時執行一次語法以確保跨網段連線時網路不會被切斷
   If tmrConnect.Interval = 0 Then
      If Now > dtNow Then
         dtNow = DateAdd("n", cntAutoQueryInterval, Now)
         ClsLawReadRstMsg 1, "select * from dual"
      End If
   End If
   'end 2024/8/7

'Added by Morgan 2016/4/14
'切回財務表單時顯示上方的工具列
If Not Me.ActiveForm Is Nothing Then
   If LCase(Left(Me.ActiveForm.Name, 6)) = "frmacc" Then
       Toolbar1.Visible = True
       StatusBar1.Visible = True
   Else
       Toolbar1.Visible = False
       StatusBar1.Visible = False
   End If
End If
'end 2016/4/14

'add by nickc 2005/05/02 電腦中心的不管
'edit by nickc 2005/09/20 加入可以印圖的控制
'If Pub_StrUserSt03 = "M51" Then Exit Sub
If Pub_StrUserSt03 = "M51" Or Pub_Can_Copy_Pic = True Then Exit Sub
    '圖檔才清
    If Clipboard.GetFormat(1) = False And Clipboard.GetFormat(2) = True And Clipboard.GetFormat(3) = False Then

        Clipboard.Clear
    End If
End Sub

'Add by Morgan 2003/12/23
'控制連線閒置超過30分鐘自動離線
Private Sub tmrConnect_Timer()

On Error GoTo ErrHnd:

   tmrConnect.Tag = tmrConnect.Tag + 1
   'Modify by Morgan 2005/2/3 改成10分鐘--薛副理
   'If tmrConnect.Tag = 30 Then
   'Modified by Morgan 2013/9/25 改回30分鐘--薛經理
   'If tmrConnect.Tag = 10 Then
   If tmrConnect.Tag = 30 Then
      Timer1.Enabled = False
      'Add by Morgan 2005/2/14
      Timer1.Interval = 0
      
'Modify by Morgan 2005/1/10 改保留原畫面不結束
'      Call CloseAllChild
'      Call SwitchMenu(False)
'2005/1/10 end

'Modify by Morgan 2005/2/16
      'cnnConnection.Close
      bolReOpen = False
      PUB_SendMailCache 'Add by Morgan 2010/6/11
      frmReopen.Show vbModal, Me
'      If MsgBox("因為閒置時間過長系統已自動離線，是否要重新連線？", vbYesNo + vbDefaultButton1 + vbExclamation) = vbYes Then
'         If PUB_ReConnect() = True Then
'            If PUB_SetUserData() = False Then
'               MsgBox "資料庫變數設定失敗！"
'            Else
'               mdiMain.bolReOpen = True
'               pub_strUserOffice = PUB_GetST06(strUserNum)
'            End If
'         End If
'      End If
 '2005/2/16 end

      If bolReOpen = True Then
         Call ReConnect
      Else
         Call mnu00_Click(1)
      End If
   End If
   
ErrHnd:

   If Err.NUMBER <> 0 Then MsgBox Err.Description
   
End Sub

Private Sub PrintReport()
   Dim rsTmp As ADODB.Recordset, intR As Integer, stSQL As String
   '若登入者為韓聖文(91028)->79075->94003->95008->95014
   'edit by nickc 2007/10/31 改成抓  table
   'If strUserNum = "95014" Then
   'edit by nickc 2007/11/05
   'If strUserNum = Pub_GetSpecMan("G") Then
   If InStr(1, Pub_GetSpecMan("G"), strUserNum) <> 0 Then
      Screen.MousePointer = vbHourglass
      PrintData1 '列印大陸案件齊備3天未完稿清單
      PrintData2 '國外新案收文3天未齊備且無關聯案件清單
      Screen.MousePointer = vbDefault
   End If
End Sub

Private Sub PrintLetter()
   Dim stTemp As String, arrNum() As String, ii As Integer
   'Added by Lydia 2016/08/31 輸入提高點數簽核主管及簽核點數
   Set Tmpfrm880004_4 = frm880004
   
   If PUB_Cache2Letter(, , False, , True) = True Then
       If MsgBox("你有報價定稿待列印，是否現在執行！", vbQuestion + vbYesNo + vbDefaultButton2, "報價定稿提醒") = vbYes Then
          mnu1601_Click 6
       End If
   End If
   
   'Added by Morgan 2015/10/12
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
   'end 2015/10/12
   
End Sub
'Add by Morgan 2010/3/19
'未發文案件查詢
Private Sub UndeliveredCaseQuery()
   
   'Modify by Morgan 2010/9/1 排除國外部程序
   'If Pub_StrUserSt03 = "P12" Then
   If Pub_StrUserSt03 = "P12" And strUserNum <> "P1003" Then
      'Modified by Lydia 2018/10/26  所有P12(內專程序)都要自動執行 (取消"專利處程序期限通知")
      'If PUB_GetST05(strUserNum) <> "73" Then
         If CheckUse("frm040210", strExec, False) = True Then
            strSql = "select 1 from executelog where el01='frm040210' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 0 Then
               Load frm040210
               '若當日第一次執行時要含已收款缺文件案件
               strSql = "select 1 from executelog where el01='frm040210' and el02='" & strUserNum & "' and el03=" & strSrvDate(1)
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 0 Then
                  frm040210.txtDoc.Text = "Y"
               End If
               'frm040210.txtDate(1) = strSrvDate(2) + 10000 '測試用
               frm040210.cmdQuery.Value = True
            End If
         End If
      'End If 'Remove by Lydia 2018/10/27
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
'            If MsgBox("是否執行  業務期限資料查詢  功能", vbYesNo, "功能！") = vbYes Then
'               pub_CallNextForm = True
'               frm100123.Show
'               frm100123.cmdSearch_Click
'            End If
'         End If
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
'               If MsgBox("是否執行  業務期限資料查詢  功能", vbYesNo, "功能！") = vbYes Then
'                  pub_CallNextForm = True
'                  frm100123.Show
'                  frm100123.cmdSearch_Click
'               End If
'            End If
'         End If
      End If
   End If
End Sub

Private Sub FcpDueCaseQurey()
Dim tmpST15 As String
   
   'Modify By Sindy 2020/8/25 Patpro:限FXX部門人員
   tmpST15 = PUB_GetStaffST15(strUserNum, 1)
   If UCase(Mid(tmpST15, 1, 1)) <> "F" Then Exit Sub
   '2020/8/25 END
   
   '電腦中心除外
   If Pub_StrUserSt03 <> "M51" Then
      'Add by Morgan 2009/12/15 改先執行FMP案已達約定期限通知 (外專程序會操作此系統)
      'Modify By Sindy 2023/9/19 取消;frm060204會抓FMP期限
'      If CheckUse("frm060206", strExec, False) = True Then
'         strSql = "select * from executelog where el01='frm060206' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'         If intI <> 1 Then
'            Load frm060206
'            frm060206.cmdQuery.Value = True
'         End If
'      'end 2009/12/15
'      Else
      If CheckUse("frm060204", strExec, False) = True Then
         strSql = "select * from executelog where el01='frm060204' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI <> 1 Then
            Load frm060204
            frm060204.cmdQuery(0).Value = True
         End If
      End If
   End If
End Sub

'Add By Sindy 2015/7/2
Public Sub SetTmpForm()
   Set Tmpfrm210147 = frm210147
   Set Tmpfrm210148 = frm210148
   Set Tmpfrm06010616 = frm06010616
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
   If m_blnActivated = False Then
      m_blnActivated = True
      '業務期限資料查詢
      SalesDueCaseQuery
      'Added by Lydia 2018/10/18 專利處程序期限通知
      'Remove by Lydia 2018/10/26 改用未發文案件查詢
      'If Pub_StrUserSt15 = "P12" Then
      '      If CheckUse("frm040211", strExec, False) = True Then
      '         strSql = "select * from executelog where el01='frm040211' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " "
      '         intI = 1
      '         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      '         If intI <> 1 Then
      '            Load frm040211
      '            frm040211.cmdQuery.Value = True
      '         End If
      '      End If
      'End If
      'end 2018/10/18
      'end 2018/10/26
      '專利處報表
      PrintReport
      '報價定稿
      PrintLetter
   End If
   
   '本查詢需考慮當閒置太久重新登入且已經是下午時須再次執行故與單獨控制
   If pub_bolInformCheck = True Then
      'Add by Morgan 2010/3/19
      '未發文案件查詢
      UndeliveredCaseQuery
      
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

Private Sub MDIForm_Load()
   'Add by Morgan 2003/12/23
   '控制連線閒置超過30分鐘自動關閉程式
   If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") = 0 Then
      Set eventConn = cnnConnection
      tmrConnect.Interval = 60000
   End If
   
   Dim lngValue, lngBufferSize As Long, intCounter As Integer
   Dim strUserId As String * 10, strLocalId As String
   Dim strSysKind As String

    '若登入成功
    If pub_str_LoginSucceeded = "1" Then
        'Add by Amy 2014/05/22 測式用
        If Pub_StrUserSt03 = "M51" Then
            mnuChUser.Visible = True
            tmrConnect.Interval = 0
        Else
            mnuChUser.Visible = False
        End If
        'end 2014/05/22
      
        Me.Timer1.Interval = 100
        'Modify By Cheng 2003/07/10
        'Begin
    '   Set cnnConnection = objPublicData.Connection
        'End
        
       strSysKind = GetSystemKindByNick
       
'       'Added by Morgan 2016/1/19 薪資查詢測試
'       If strUserNum = "94007" Or strUserNum = "71011" Then
'         mnu23(6).Visible = True
'       Else
'         mnu23(6).Visible = False
'       End If
'       'end
       
       'add by nickc 2006/06/09 可以查詢維護紀錄
       If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Or Pub_StrUserSt03 = "M51" Then
            mnuDML(0).Visible = True
       Else
            mnuDML(0).Visible = False
       End If
       
       If bolFNation = False Then
    'Ken 90/07/06
    '      mnu10(14).Visible = False
    '      mnu10(15).Caption = "以國籍查詢申請人"
          mnu101(10).Visible = False 'Modify By Amy 2014/05/05 Sindy 2011/10/3
          mnu102(7).Visible = False 'Modify By Sindy 2011/10/3
    'Ken 90/07/06
          mnu101(6).Visible = False 'Modify By Amy 2014/05/05 Sindy 2011/10/3
          '92.3.17 add by sonia
          mnu102(2).Visible = False 'Modify By Sindy 2011/10/3
       End If
'        'add by nickc 2007/10/02 林淑真、杜副總、電腦中心 才秀的
'        '2009/8/12 MODIFY BY SONIA 加入68001江總
'        'modify by sonia 2014/6/9 +美珍77027
'        '2015/7/24 modify by sonia +林總94007,何主秘68009,-江總68001-小真65001
'        If strUserNum = "94007" Or strUserNum = "77027" Or strUserNum = "68006" Or strUserNum = "68009" Or Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
'          mnu2103(5).Visible = True
'          mnu2103(6).Visible = True
'          'mnu2103(7).Visible = True
'          mnu2103(13).Visible = True 'Add by Amy 2015/08/05
'        Else
'          mnu2103(5).Visible = False
'          mnu2103(6).Visible = False
'          'mnu2103(7).Visible = False
'          mnu2103(13).Visible = False 'Add by Amy 2015/08/05
'        End If
'
'        'Add By Sindy 2010/11/4
'        '74028.邱素蓮、杜副總、電腦中心 才秀的
'        '2015/7/24 modify by sonia +林總94007,何主秘68009
'        If strUserNum = "74028" Or strUserNum = "68006" Or strUserNum = "94007" Or strUserNum = "68009" Or Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
'            mnu2103(7).Visible = True
'        Else
'            mnu2103(7).Visible = False
'        End If
'
'        If Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
'            mnu2103(3).Visible = True
'        '2012/2/13 add by sonia 加林特助94007權限
'        ElseIf strUserNum = "94007" Then
'            mnu2103(3).Visible = True
'        '2012/2/13 end
'        Else
'            mnu2103(3).Visible = False
'        End If
        
'        'Add by Sindy 2014/4/18 陳淑芳87027及電腦中心才看的到電話分機維護
'        If Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Or strUserNum = "87027" Then
'            mnu23(5).Visible = True
'        Else
'            mnu23(5).Visible = False
'        End If
      
      'Added by Morgan 2013/9/12
      If Left(Pub_StrUserSt03, 2) = "P1" Then
         mnu23(1).Caption = "會議室/檢索系統預約作業"
      End If
      'end 2013/9/12
      
      'Add by Amy 2017/01/25
      If strSrvDate(1) >= 20170202 Then
        mnu23(8).Visible = True
      Else
        mnu23(8).Visible = False
      End If
      
      'Add by Morgan 2009/4/29
      If Left(Pub_StrUserSt15, 1) = "F" Then
         mnuTitle(9).Visible = False
      End If
      
      '2009/12/22 add by sonia M51及王副總才可以看到
      'If Pub_StrUserSt03 = "M51" Or strUserNum = "71011" Then
      'If Pub_StrUserSt03 = "M51" Or CheckUse("frm050207", strExec) = True Then
      If Pub_StrUserSt03 = "M51" Then
          mnu0402(7).Visible = True
      Else
          mnu0402(7).Visible = False
      End If
      '2009/12/22 end
      'Added by Lydia 2022/07/01 客戶提供文件處理：給外專人員操作寰華案
      If Pub_StrUserSt03 = "M51" Or Left(Pub_StrUserSt15, 2) = "F2" Then
          mnu0401(24).Visible = True
      Else
          mnu0401(24).Visible = False
      End If
      'end 2022/07/01
      
      'Add by Amy 2018/08/09 待處理區
      If Val(strSrvDate(1)) < 非P結案電子化啟用日 Then
        mnu0501(12).Visible = False
      End If
      'end 2018/08/09
      
      'Added by Morgan 2018/6/29
      If Val(strSrvDate(1)) < CFP第一階段電子化啟用日 Then
         mnu0501(13).Visible = False
      End If
      'end 2018/6/29
      
      'Added by Morgan 2018/8/17
      If Val(strSrvDate(1)) < CFP指示信電子化啟用日 Then
         mnu0501(14).Visible = False
      End If
      'end 2018/8/17
      
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
'edit by nickc 2007/02/02 不用 dll 了
'Set objPublicData = Nothing

   ' 90.08.16 modify by louis (釋放Word物件)
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

Private Sub mnu0401_Click(Index As Integer)
ToolHide
intPCaseKind = 專利
intPWhere = 國內
Select Case Index
   Case 1   '內專分案
      If CheckUse("frm040101", strExec) = True Then
         frm040101.Show
      End If
   Case 2   '各式申請書 'Memo by Lydia 2023/09/22 frm040110由此進入，改名為「各式申請書／書表」
      If CheckUse("frm040103_1", strExec) = True Then
         frm040103_1.Show
      End If
   'Add By Cheng 2002/07/09
   Case 3   '國外指示信
      If CheckUse("frm040104_1", strExec) = True Then
         frm040106_1.Show
      End If
   Case 4   '內專發文
      If CheckUse("frm040104_1", strExec) = True Then
         frm040104_1.Show
      End If
   Case 5   '申請案號輸入
      If CheckUse("frm04010401", strExec) = True Then
         frm04010401.Show
      End If
      
'Removed by Morgan 2015/9/10 移到
'   Case 8   '國內外案件資料維護
'      If CheckUse("frm050106_1", strExec) = True Then
'         frm050106_1.intWhereToGo = 0
'         frm050106_1.Show
'      End If
'   Case 9   '大陸發明案件資料維護
'      If CheckUse("frm040108_1", strExec) = True Then
'         frm040108_1.Show
'      End If
'   'Add by Morgan 2004/6/14   '一案兩申請案件資料維護
'   Case 10  '一案兩申請案件資料維護
'      If CheckUse("frm040109_1", strExec) = True Then
'         frm040109_1.Show
'      End If
'   'add by nickc 2005/06/23 大陸香港案
'   Case 11  '大陸香港案件資料維護
'      If CheckUse("frm050109_1", strExec) = True Then
'         frm050109_1.intWhereToGo = 0
'         'frm050109_1.iK_CM10 = "4" 'Added by Lydia 2015/07/28 未正式上線
'         frm050109_1.Show
'      End If
'
'   'Added by Lydia 2015/07/28
'    Case 12 '大陸澳門案
'      If CheckUse("frm050109_1", strExec) = True Then
'         frm050109_1.intWhereToGo = 0
'         'frm050109_1.iK_CM10 = "5" 未正式上線
'         frm050109_1.Show
'      End If
'end 2015/9/10
   
   'Add by Morgan 2010/12/14   '領證年費整批發文
    'Modified by Lydia 2015/07/28 大陸澳門案以後,功能表index + 1
   'Case 12
    Case 13
      If CheckUse("frm040104_i", strExec) = True Then
         'Added by Morgan 2013/7/29
         If PUB_CheckFormExist("frm040104_7") Then
            MsgBox "【內專領證年費整批發文】不可與【" & frm040104_7.Caption & "】同時執行！", vbExclamation
         ElseIf PUB_CheckFormExist("frm040104_a") Then
            MsgBox "【內專領證年費整批發文】不可與【" & frm040104_a.Caption & "】同時執行！", vbExclamation
         Else
         'end 2013/7/29
         
            frm040104_i.Show
         End If
         
         
      End If
      
   'Add by Sindy 2011/01/03   '解聘書
   'Modified by Lydia 2015/07/28 大陸澳門案以後,功能表index + 1
   'Case 13
   'Mark by Lydia 2023/09/22  併入各式申請書frm040103_1; 功能表已刪除項目
   'Case 14
   '   If CheckUse("frm040110", strExec) = True Then
   '      frm040110.Show
   '   End If
   'end 2023/09/22
'Remove by Morgan 2010/11/9 表單已滿，目前沒用的功能先移除
'   'Add by Morgan 2007/6/23
'   Case 12 '委任通知函客戶回覆單輸入
'      If CheckUse("frm04010701", strExec) = True Then
'         frm04010701.Show
'      End If
'
'   'Add by Morgan 2007/6/30
'   Case 13 '委任通知函客戶清單確認
'         If CheckUse("frm04010702", strExec) = True Then
'            frm04010702.Show
'         End If
'
'   '2007/7/3 ADD BY SONIA
'   Case 14 '重新委任發文(不印申請書)
'         If CheckUse("frm04010703", strExec) = True Then
'            frm04010703.Show
'         End If
'
'   'Add by Morgan 2007/7/3
'   Case 15 '重新委任批次收文(未回覆客戶)
'         If CheckUse("frm04010704", strExec) = True Then
'            frm04010704.Show
'         End If
'end 2010/11/9
   'Add By Sindy 2013/5/17
   'Modified by Lydia 2015/07/28 大陸澳門案以後,功能表index + 1
   'Case 14 '待送件區
   Case 15
      If CheckUse("frm090202_4", strExec) = True Then
         frm090202_4.m_ProState = "P"
         frm090202_4.Show
      End If
      
   'Add By Sindy 2013/10/3
   'Modified by Lydia 2015/07/28 大陸澳門案以後,功能表index + 1
   'Case 15 '電子送件電子檔整批匯入
   Case 16
      If CheckUse("frm040111", strExec) = True Then
         frm040111.Show
      End If
   'Added by Morgan 2018/4/24 改回第一層,取消期限管制輸入及查詢功能-跟郭確認過
   Case 17 '電子公文來函
      If CheckUse("frm04010516", strExec) = True Then
         frm04010516.Show
      End If
   'end 2018/4/24
   
   'Added by Morgan 2014/4/18
   'Modified by Lydia 2015/07/28 大陸澳門案以後,功能表index + 1
   'Case 17 '公文來函文檔整批匯入
   Case 18
      If CheckUse("frm040112", strExec) = True Then
         frm040112.Show
      End If
      
   'Add By Sindy 2014/7/21
   'Modified by Lydia 2015/07/28 大陸澳門案以後,功能表index + 1
   'Case 18 '非台灣案發文後暫緩
   Case 19
      If CheckUse("frm040114_1", strExec) = True Then
         frm040114_1.Show
      End If
      
   'Add By Amy 2014/08/06
   'Modified by Lydia 2015/07/28 大陸澳門案以後,功能表index + 1
   'Case 19 '收據整批匯入
   Case 20
      If CheckUse("frm040115", strExec) = True Then
         frm040115.Show
      End If
   
   'Added by Morgan 2016/6/2
   Case 21 '代理人來函匯入
      If CheckUse("frm040120", strExec) = True Then
         frm040120.Show
      End If
      
   'Added by Morgan 2017/2/17
   Case 22 '電子收據匯入
      If CheckUse("frm040121", strExec) = True Then
         frm040121.Show
      End If
      
   'Add By Amy 2014/08/06
   'Modified by Lydia 2015/07/28 大陸澳門案以後,功能表index + 1
   'Case 20 '待處理區
   'Modified by Morgan 2016/6/2 index+1
   'Modified by Morgan 2017/2/17 index+1
   Case 23
      If CheckUse("frm210149", strExec) = True Then
         frm210149.m_ProState = "P"
         frm210149.Show
      End If
    'Added by Lydia 2022/06/28
    Case 24 '客戶提供文件處理
        If CheckUse("frm060121", strExec) = True Then
            frm060121.Text1 = "P"
            frm060121.Text1.Locked = True
            frm060121.Show
        End If
    End Select
End Sub

'Add By Sindy 2018/5/2
Public Sub frm090202_4CallFrm(strSendFrmType As String, m_PA11 As String, _
   m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String, _
   m_EEP01 As String)
Dim m_SendRecvForm As Form '發文作業
   Select Case strSendFrmType
      Case "P"
         Set m_SendRecvForm = frm040104_1
         m_SendRecvForm.Show
         
         'Removed by Morgan 2019/9/11 待送件區發文改都用本所案號(同申請號可能有多個舉發案 Ex:200710186670.7 P122180,P116700)
         'If m_PA11 <> "" Then
         '   m_SendRecvForm.Option1(1).Value = True
         '   m_SendRecvForm.Text5 = m_PA11 'm_EEP01
         'Else
         'end 2019/9/11
         
            m_SendRecvForm.Option1(0).Value = True
            m_SendRecvForm.Text1 = m_CP01
            m_SendRecvForm.Text2 = m_CP02
            m_SendRecvForm.Text3 = m_CP03
            m_SendRecvForm.Text4 = m_CP04
         'End If 'Removed by Morgan 2019/9/11
         m_SendRecvForm.bolIsEMPFlow = True
         m_SendRecvForm.Command1_Click
         Set m_SendRecvForm = Nothing
      Case "CFP"
         Set m_SendRecvForm = frm050102_1
         m_SendRecvForm.Show
         m_SendRecvForm.OptChoose(0).Value = True
         m_SendRecvForm.txtReceiveCode = m_EEP01
         m_SendRecvForm.bolIsEMPFlow = True
         Call m_SendRecvForm.cmdOK_Click(2)
         Call m_SendRecvForm.cmdOK_Click(0)
         Set m_SendRecvForm = Nothing
   End Select
End Sub

Private Sub mnu040105_Click(Index As Integer)
ToolHide
intPCaseKind = 專利
intPWhere = 國內
Select Case Index
   Case 1   '實審通知日輸入
      If CheckUse("frm04010501", strExec) = True Then
         frm04010501.Show
      End If
   'Add by Morgan 2009/11/24
   Case 2   '初審及公佈通知來函輸入
      If CheckUse("frm04010514_1", strExec) = True Then
         frm04010514_1.Show
      End If
   Case 3   '核准函輸入
      If CheckUse("frm04010502_1", strExec) = True Then
         frm04010502_1.Show
      End If
   Case 4  '核駁函輸入
      If CheckUse("frm04010503_1", strExec) = True Then
         frm04010503_1.Show
      End If
'   Case 5   '通知領證函輸入
'      If CheckUse("frm04010510_1", strExec) = True Then
'         frm04010510_1.Show
'      End If
   Case 6   'Memo by Lydia 2024/01/17 「消滅函輸入」更名為「消滅函／視為撤回輸入」
      If CheckUse("frm04010511_1", strExec) = True Then
         frm04010511_1.Show
      End If
   Case 7   '一般來函輸入
      If CheckUse("frm04010504_1", strExec) = True Then
         frm04010504_1.Show
      End If
   Case 8   '證書號數輸入
      If CheckUse("frm04010505_1", strExec) = True Then
         frm04010505_1.Show
      End If
   Case 9   '異議/舉發受理函輸入
      If CheckUse("frm04010506_1", strExec) = True Then
         frm04010506_1.Show
      End If
      
   'Added by Morgan 2020/1/16 從通知函移來
   Case 10   '年費逾期補繳通知函輸入
      If CheckUse("frm040324", strExec) = True Then
         frm040324.iKind = 1
         frm040324.Show
      End If
      
   Case 11   '實審請求期限屆滿前通知函輸入
      If CheckUse("frm040324", strExec) = True Then
         frm040324.iKind = 2
         frm040324.Show
      End If
   'end 2020/1/16
   
End Select
End Sub
Private Sub mnu040106_Click(Index As Integer)
ToolHide
intPCaseKind = 專利
intPWhere = 國內
Select Case Index
   Case 1   '代理人已收達/己提申
      If CheckUse("frm04010507_1", strExec) = True Then
         frm04010507_1.Show
      End If
   Case 2   '代理人通知修正
      If CheckUse("frm04010508_1", strExec) = True Then
         frm04010508_1.Show
      End If
   Case 3   '所外鑑定報告結果
      If CheckUse("frm04010509_1", strExec) = True Then
         frm04010509_1.Show
      End If
   Case 4   '其他來函輸入
      If CheckUse("frm02010603_1", strExec) = True Then
         frm02010603_1.Show
         frm02010603_1.Caption = "其他來函輸入"
      End If
    Case 5   '代理人信件收達管制
      If CheckUse("frm04010512", strExec) = True Then
         StrStartSystemByNick = "P"           '2008/8/29 add by sonia 依系統別預設系統類別
         frm04010512.Show
      End If
    'Add by Morgan 2009/11/17
    Case 6   '已順稿
      If CheckUse("frm04010513", strExec) = True Then
         frm04010513.Show
      End If
End Select
End Sub

Private Sub mnu0402_Click(Index As Integer)
   ToolHide
   Select Case Index
      '代理人新案案件統計
      Case 1
         If CheckUse("frm050201", strExec) = True Then
            StrStartSystemByNick = "P,PS"
            frm050201.Show
         End If
'edit by nickc 2005/07/22
'      Case 2
'         If CheckUse("frm040202", strExec) = True Then
'            frm040202.Show
'         End If
      '未請款明細查詢
      Case 3
         If CheckUse("frm050203", strExec) = True Then
            StrStartSystemByNick = "P,PS"
            frm050203.Show
         End If
      '審查委員准駁統計
      Case 4
         If CheckUse("frm040204", strExec) = True Then
            StrStartSystemByNick = "CFP,P,FCP"
            frm040204.Show
         End If
      'FC收款請款點數查詢
      Case 5
         If CheckUse("frm040205", strExec) = True Then
            StrStartSystemByNick = GetSystemKindByNick
            frm040205.Show
         End If
      '代理人案件性質統計
      Case 6
         If CheckUse("frm050204_1", strExec) = True Then
            StrStartSystemByNick = GetSystemKindByNick
            frm050204_1.Show
         End If
      '員工查詢印表記錄檔查詢   2009/12/22 add by sonia
      Case 7
         If CheckUse("frm050207", strExec) = True Then
            StrStartSystemByNick = GetSystemKindByNick
            frm050207.Show
         End If
      '未發文案件查詢 Add by Morgan 2010/3/18
      Case 8
         If CheckUse("frm040210", strExec) = True Then
            StrStartSystemByNick = GetSystemKindByNick
            frm040210.Show
         End If
      'Add by Morgan 2011/4/20
      Case 9 '接洽記錄單查詢及列印
         'Memo by Lydia 2021/05/18 更名為「自動收文接洽單查詢/列印」
         'Modify By Sindy 2023/1/6 更名為「電子收文接洽單查詢」
         frm12040152.Show
         
      'Added by Lydia 2018/10/18
      'Remove by Lydia 2018/10/26 改用未發文案件查詢
      'Case 11 '專利處程序期限通知
      '   If CheckUse("frm040211", strExec) = True Then
      '      frm040211.Show
      '   End If
      'end 2018/10/26
      Case Else
   End Select
End Sub

Private Sub mnu0403_Click(Index As Integer)
   intPCaseKind = 專利
   intPWhere = 國內
   ToolHide
   Select Case Index
'Remove by Morgan 2005/8/25 移到 mnu040301
'        'Add By Cheng 2003/05/21
'        '公開通知函
'      Case 1
'         If CheckUse("frm040325", strExec) = True Then
'            frm040325.Show
'         End If
'      Case 2
'         If CheckUse("frm040301", strExec) = True Then
'            frm040301.Show
'         End If
'      Case 3
'         If CheckUse("frm040302", strExec) = True Then
'            frm040302.Show
'         End If
'      Case 4
'         If CheckUse("frm040303", strExec) = True Then
'            frm040303.Show
'         End If
'      '92.3.5 Add By sonia
'      Case 5 '年費逾期補繳通知函
'         If CheckUse("frm040324", strExec) = True Then
'            frm040324.Show
'         End If
'      '92.3.5 end
'      'Add By Cheng 2002/06/24
'      Case 6 '其他通知函/聯絡單
'         If CheckUse("frm040322", strExec) = True Then
'            frm040322.Show
'         End If

'Removed by Morgan 2020/4/15
'      'Add By Cheng 2002/06/24
'      Case 7   '領證遭異資料還原作業
'         If CheckUse("frm040323", strExec) = True Then
'            frm040323.Show
'         End If

      Case 8   '期限管制表
         If CheckUse("frm040304", strExec) = True Then
            frm040304.Show
         End If
      Case 9   '代理人案件收達/提申管制表
         If CheckUse("frm050303", strExec) = True Then
            frm050303.Show
         End If
      Case 10  '收文未發文明細表
         If CheckUse("frm050304", strExec) = True Then
            frm050304.Show
         End If
      Case 11  '催審函/催審表
         If CheckUse("frm040307", strExec) = True Then
            frm040307.Show
         End If
      Case 12  '智權人員收文明細表
         If CheckUse("frm040308", strExec) = True Then
            frm040308.Show
         End If
      Case 13  '收文簿
         If CheckUse("frm050307", strExec) = True Then
            StrStartSystemByNick = "P"
            frm050307.Show
         End If
      Case 14  '發文簿
         If CheckUse("frm040310", strExec) = True Then
            frm040310.Show
         End If
      Case 15  '核准(駁)簿
         If CheckUse("frm040311", strExec) = True Then
            frm040311.Show
         End If
      Case 16  '大陸發明案參考資料表
         If CheckUse("frm040312", strExec) = True Then
            frm040312.Show
         End If
      Case 17  '顧問客戶委辦案件明細表
         If CheckUse("frm040313", strExec) = True Then
            frm040313.Show
         End If
      Case 18  '後金案件表
         If CheckUse("frm050314", strExec) = True Then
            frm050314.Show
         End If
      Case 19  '延期明細表
         If CheckUse("frm050315", strExec) = True Then
            frm050315.Show
         End If
      Case 20  '不出名案件明細表
         If CheckUse("frm040316", strExec) = True Then
            frm040316.Show
         End If
      Case 21  '代理人案件總簿
         If CheckUse("frm050316", strExec) = True Then
            frm050316.Show
         End If
      Case 22  '客戶案件總簿
         If CheckUse("frm050317", strExec) = True Then
            'Modify by Amy 2017/08/04 原:frm050317
            frm0503171.Show
         End If
      Case 23  '代理人/申請人名單
         If CheckUse("frm050318", strExec) = True Then
            frm050318.Show
         End If
      Case 25  '地址條列印
         If CheckUse("frm083014", strExec) = True Then
            frm083014.Show
         End If
      'Add by Morgan 2005/8/25
      Case 26  '核准領證期限表
         If CheckUse("frm040328", strExec) = True Then
            frm040328.Show
         End If
      'Add by Morgan 2005/6/28
      Case 27  '智慧局年費通知核對清單
         If CheckUse("frm040331", strExec) = True Then
            frm040331.Show
         End If
      'Add by Sindy 2011/12/27
      Case 28  '證書PDF列印
         If CheckUse("frm040334", strExec) = True Then
            frm040334.Show
         End If
      'Add by Amy 2021/09/10
      Case 29 '資策會專利案件季報表
        If CheckUse("frm040336", strExec) = True Then
            frm040336.Show
         End If
      'Added by Lydia 2022/04/12
      Case 30 '資策會收到證書清單
        If CheckUse("frm040337", strExec) = True Then
            frm040337.Show
         End If
   End Select
End Sub

Private Sub mnu0404_Click(Index As Integer)
   intPCaseKind = 1
   intPWhere = 0
   ToolHide
   Select Case Index
      Case 1   '收文統計表
         If CheckUse("frm040401", strExec) = True Then
            frm040401.Show
         End If
      Case 2   '發文統計表
         If CheckUse("frm040402", strExec) = True Then
            frm040402.Show
         End If
      Case 3   '准駁預估統計表
         If CheckUse("frm040403", strExec) = True Then
            frm040403.Show
         End If
      Case 4   '准駁統計總表
         If CheckUse("frm040404", strExec) = True Then
            frm040404.Show
         End If
      Case 5   '准駁統計明細表
         If CheckUse("frm040405", strExec) = True Then
            frm040405.Show
         End If
      Case 6   '代理人新案案件統計表
         If CheckUse("frm050404", strExec) = True Then
            frm050404.Show
         End If
        'Add By Cheng 2003/12/03
      Case 7   '代理人案件年度統計表
         If CheckUse("frm050407", strExec) = True Then
            frm050407.Show
         End If
        'End
      Case 8   '代理人/申請人新申請案排行榜
         If CheckUse("frm050405", strExec) = True Then
            StrStartSystemByNick = "P"
            frm050405.Show
         End If
      Case 9   '逾期未結案統計表
         If CheckUse("frm084004", strExec) = True Then
            frm084004.Tag = 4
            frm084004.Show
         End If
   End Select
End Sub

Private Sub mnu0405_Click(Index As Integer)
Dim strSysID As String 'Add By Sindy 2014/7/4
   
   ToolHide
   strSysKind = "1"
   Select Case Index
      Case 1   '專利案件基本資料維護
         If CheckUse("frm050701", strExec) = True Then
            strSysKind = "P"
            frm050701.Show
         End If
      Case 2   '服務業務基本資料維護
         If CheckUse("frm050702", strExec) = True Then
            strSysKind = "PS"
            frm050702.Show
         End If
      Case 3   '案件進度檔資料維護
         If CheckUse("frm075004_1", strExec) = True Then
            strSysKind = "P"
            frm075004_1.Show
         End If
      Case 4   '下一程序資料
         If CheckUse("frm075007_1", strExec) = True Then
            strSysKind = "P"
            frm075007_1.Show
         End If
      Case 5   '國外代理人資料
         If CheckUse("frm050705", strExec) = True Then
            strSysKind = "P"
            frm050705.Show
         End If
      Case 6   '變更事項
         If CheckUse("frm050706", strExec) = True Then
            strSysKind = "P"
            frm050706.Show
         End If
      Case 7   '延期記錄資料維護
         If CheckUse("frm050707", strExec) = True Then
            strSysKind = "P"
            frm050707.Show
         End If
      Case 8   '案件國家收費表維護
         If CheckUse("frm12040102", strExec) = True Then
            strSysKind = "P"
            frm12040102.Show
         End If
      Case 9  '客戶發明人資料維護
         If CheckUse("frm050709", strExec) = True Then
            strSysKind = "P"
            frm050709.Show
         End If
      Case 11  '代理人變更名稱作業
         If CheckUse("frm140103", strExec) = True Then
            frm140103.Show
         End If
      'add by nick 2004/07/14
      Case 12  '客戶減免身分維護
         If CheckUse("frm050715", strExec) = True Then
            frm050715.Show
         End If
      'add by nickc 2007/10/17
      Case 13  '系統特殊設定
         If CheckUse("frm050716", strExec) = True Then
            frm050716.Show
         End If
      'Add by Morgan 2008/5/28
      Case 14  '大陸領證報價資料維護 '2015/01/08 不限大陸,更名為領證報價資料維護
         If CheckUse("frm12040115", strExec) = True Then
            frm12040115.Show
         End If
      'Add by Lydia 2015/01/05
      Case 15  '年費報價資料維護
         If CheckUse("frm12040116", strExec) = True Then
            frm12040116.Show
         End If
      'Add by Morgan 2008/11/19
      Case 16  '依案件性質設定各國催審提申期限
         If CheckUse("frm12040102_1", strExec) = True Then
            frm12040102_1.Show
         End If
      'Add by Morgan 2011/12/13
      Case 17  '電子報排程維護
         If CheckUse("frm140410_1", strExec) = True Then
            frm140410_1.Show
         End If
      'Add By Sindy 2014/7/9
      Case 18 '台灣專利總委任書正本案號維護
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
      'Add By Sindy 2014/10/27
      Case 19 '更換FC代理人作業
         If CheckUse("frm110104_1", strExec) = True Then
            frm110104_1.Show
         End If
      'Added by Lydia 2016/10/24
      Case 20 '申請人指定國外代理人維護
         If CheckUse("frm050718", strExec) = True Then
            frm050718.Show
         End If
      'Added by Morgan 2018/7/23
      Case 21 '程序人員核判表維護
         If CheckUse("frm050719", strExec) = True Then
            frm050719.Show
         End If
   End Select
End Sub

Private Sub mnu040601_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1   '專利公報資料維護
         If CheckUse("frm04060101_1", strExec) = True Then
            frm04060101_1.Show
         End If
      Case 2   '市場佔有率統計表列印
         If CheckUse("frm04060102", strExec) = True Then
            frm04060102.Show
         End If
      Case 3   '專利公報查詢列印
         If CheckUse("frm04060103_1", strExec) = True Then
            frm04060103_1.Show
         End If
      Case 4   '國內代理人名稱查詢
         If CheckUse("frm04060104", strExec) = True Then
            frm04060104.Show
         End If
      Case 5   '國內市場佔有率查詢
         If CheckUse("frm04060105_1", strExec) = True Then
            frm04060105_1.Show
         End If
      Case 6   '國內公報代理人資料維護
         If CheckUse("frm04060106", strExec) = True Then
            frm04060106.Show
         End If
      Case 7   '國內公報代理人換事務所作業
         If CheckUse("frm04060107", strExec) = True Then
            frm04060107.Show
         End If
      'Add By Sindy 2011/5/30
      Case 8   '國內公報代理人資料列印
         If CheckUse("frm04060109", strExec) = True Then
            frm04060109.strTA01 = "P"
            frm04060109.Show
         End If
      'Add By Sindy 2011/12/14
      Case 9   '專利公報轉檔作業
         If CheckUse("frm04060110", strExec) = True Then
            frm04060110.Show
         End If
      'Add By Sindy 2017/5/10
      Case 10   '公報特殊字對照檔
         If CheckUse("frm030617", strExec) = True Then
            frm030617.Show
            frm030617.Caption = "公報特殊字對照檔"
            frm030617.SSTab1.TabEnabled(0) = False
            frm030617.SSTab1.Tab = 1 'Add By Sindy 2018/5/9
         End If
   End Select
End Sub

Private Sub mnu040602_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1   '大陸專利公報資料維護
         If CheckUse("frm04060201_1", strExec) = True Then
            frm04060201_1.Show
         End If
      Case 2   '大陸專利市場佔有率統計表
         If CheckUse("frm04060202", strExec) = True Then
            frm04060202.Show
         End If
      Case 3   '大陸專利公報查詢列印
         If CheckUse("frm04060203_1", strExec) = True Then
            frm04060203_1.Show
         End If
      Case 4   '大陸公報開拓函列印
         If CheckUse("frm04060204", strExec) = True Then
            frm04060204.Show
         End If
      Case 5   '大陸事務所資料維護
         If CheckUse("frm04060205", strExec) = True Then
            frm04060205.Show
         End If
      Case 6   '大陸開拓客戶資料維護
         If CheckUse("frm04060206", strExec) = True Then
            frm04060206.Show
         End If
   End Select
End Sub

Private Sub mnu040603_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1   '專利公開公報資料維護
         If CheckUse("frm04060301_1", strExec) = True Then
            frm04060301_1.Show
         End If
      Case 2   '國內公開後實審輸入
         If CheckUse("frm04060304_1", strExec) = True Then
            frm04060304_1.Show
         End If
      Case 3   '公開市場佔有率統計表列印
         If CheckUse("frm04060302", strExec) = True Then
            frm04060302.Show
         End If
      Case 4   '專利公開公報查詢列印
         If CheckUse("frm04060303_1", strExec) = True Then
            frm04060303_1.Show
         End If
      Case 5   '國內公開市場佔有率查詢
         If CheckUse("frm04060305_1", strExec) = True Then
            frm04060305_1.Show
         End If
      'Add By Sindy 2011/12/20
      Case 6   '專利公開公報轉檔作業
         If CheckUse("frm04060306", strExec) = True Then
            frm04060306.Show
         End If
   End Select
End Sub

Private Sub mnu0501_Click(Index As Integer)
   ToolHide
   intPCaseKind = 專利
   intPWhere = 國外_CF
   Select Case Index
      Case 1   '分案
         If CheckUse("frm050101_1", strExec) = True Then
            frm050101_1.Show
         End If
      Case 2   'CFP發文
         If CheckUse("frm050102_1", strExec) = True Then
            frm050102_1.Show
         End If
      Case 3   '代理人案件提申
         If CheckUse("frm050103_1", strExec) = True Then
            frm050103_1.Show
         End If
      Case 6   '國內外案件資料維護
         If CheckUse("frm050106_1", strExec) = True Then
            frm050106_1.intWhereToGo = 0
            frm050106_1.Show
         End If
      Case 7   '美國IDS資料對照維護
         If CheckUse("frm050107_1", strExec) = True Then
            frm050107_1.Show
         End If
      'Add By Cheng 2002/05/14
      Case 8   '國內外案件刪除作業
         If CheckUse("frm050108", strExec) = True Then
            frm050108.Show
         End If
      'Add by Morgan 2004/6/14
      Case 9   '一案兩申請案件資料維護
         If CheckUse("frm040109_1", strExec) = True Then
            frm040109_1.Show
         End If
      '2011/5/20 ADD BY SONIA
      Case 10  'CFP申請文件齊備維護
         If CheckUse("frm040209", strExec) = True Then
            frm040209.Show
         End If
      'Add By Sindy 2013/5/17
      Case 11 '待送件區
         If CheckUse("frm090202_4", strExec) = True Then
            frm090202_4.m_ProState = "CFP"
            frm090202_4.Show
         End If
      'Add by Amy 2018/06/27
      Case 12 '待處理區
        If CheckUse("frm210149", strExec) = True Then
           frm210149.m_ProState = "CFP"
           frm210149.Show
        End If
        
      'Added by Morgan 2018/6/29
      Case 13 '代理人來函匯入
         If CheckUse("frm040120", strExec) = True Then
            frm040120.Show
         End If
         
      'Added by Morgan 2018/8/17
      Case 14 '指示信判發作業
         If CheckUse("frm040119CFP", strExec) = True Then
            frm040119.m_ProState = "CFP"
            frm040119.Show
         End If
         
      'Added by Morgan 2025/11/4
      Case 15 '外翻人員給案維護
         If CheckUse("frm050110", strExec) = True Then
            frm050110.Show
         End If
   End Select
End Sub
Private Sub mnu050104_Click(Index As Integer)
   ToolHide
   intPCaseKind = 專利
   intPWhere = 國外_CF
   Select Case Index
      Case 1   '一般來函
         If CheckUse("frm05010401_1", strExec) = True Then
            frm05010401_1.Show
            frm05010401_1.Caption = "一般來函輸入"
         End If
      Case 2   '公開公告資料輸入
         If CheckUse("frm05010402_1", strExec) = True Then
            frm05010402_1.intChoose = 1
            frm05010402_1.Caption = "公開公告資料輸入"
         End If
      Case 3   '證書號數輸入
         If CheckUse("frm05010402_1", strExec) = True Then
            frm05010402_1.intChoose = 2
            frm05010402_1.Caption = "證書號數輸入"
         End If
      Case 4   '消滅函輸入
         If CheckUse("frm05010404_1", strExec) = True Then
            frm05010404_1.Show
            frm05010404_1.Caption = "消滅函輸入"
         End If
      'Add by Morgan 2008/12/15
      Case 5   '年費逾期補繳通知函
         If CheckUse("frm05010405_1", strExec) = True Then
            frm05010405_1.Show
         End If
      'Add by Morgan 2010/6/15
      Case 6   '實體審查、領證費逾期補繳通知函
         If CheckUse("frm05010406_1", strExec) = True Then
            frm05010406_1.Show
         End If
   End Select
End Sub
Private Sub mnu050105_Click(Index As Integer)
   ToolHide
   intPCaseKind = 專利
   intPWhere = 國外_CF
   Select Case Index
      Case 1   '代理人已收達/已提申
         If CheckUse("frm02010601_1", strExec) = True Then
            frm02010601_1.Show
         End If
      Case 2   '代理人通知修正
         If CheckUse("frm02010602_1", strExec) = True Then
            frm02010602_1.Show
         End If
      Case 3   '代理人其他來函輸入
         If CheckUse("frm02010603_1", strExec) = True Then
            frm02010603_1.Show
            frm02010603_1.Caption = "其他來函輸入"
         End If
      Case 4   '代理人信件收達管制
         If CheckUse("frm04010512", strExec) = True Then
            StrStartSystemByNick = "CFP"           '2008/8/29 add by sonia 依系統別預設系統類別
            frm04010512.Show
         End If
      Case 3
   End Select
End Sub

Private Sub mnu0502_Click(Index As Integer)
   ToolHide
   Select Case Index
      Case 1   '代理人新案案件統計
         If CheckUse("frm050201", strExec) = True Then
            StrStartSystemByNick = "CFP,CPS"
            frm050201.Show
         End If
'edit by nickc 2005/07/22
'      Case 2
'         If CheckUse("frm040202", strExec) = True Then
'            frm040202.Show
'         End If
      Case 3   '未請款明細查詢
         If CheckUse("frm050203", strExec) = True Then
            StrStartSystemByNick = "CFP,CPS"
            frm050203.Show
         End If
      'Add By Cheng 2002/09/24
      Case 4   '代理人案件性質統計
         If CheckUse("frm050204_1", strExec) = True Then
            StrStartSystemByNick = "CFP,CPS"
            frm050204_1.Show
         End If
      'Add by Morgan 2006/8/15
'      Case 5   '未收款無法發文案件統計
'         If CheckUse("frm050205", strExec) = True Then
'            StrStartSystemByNick = "CFP,CPS"
'            frm050205.Show
'         End If
      'Add by Morgan 2008/4/15
      Case 6   '互惠代理人目標給案未輸入明細表
         If CheckUse("frm050206", strExec) = True Then
            frm050206.Show
         End If
      'Add by Morgan 2010/3/19
      Case 7   '未發文案件查詢
         If CheckUse("frm040210", strExec) = True Then
            StrStartSystemByNick = GetSystemKindByNick
            frm040210.Show
         End If
      'Add by Sindy 2012/12/19
      Case 8   'CF代理人報價附件查詢
         If CheckUse("frm050208", strExec) = True Then
            frm050208.Show
         End If
      'Added by Lydia 2018/10/18
      'Remove by Lydia 2018/10/26 改用未發文案件查詢
      'Case 9 '專利處程序期限通知
      '   If CheckUse("frm040211", strExec) = True Then
      '      frm040211.Show
      '   End If
      'end 2018/10/26
   End Select
End Sub

Private Sub mnu0503_Click(Index As Integer)
    intPCaseKind = 專利
    intPWhere = 國外_CF
    ToolHide
    Select Case Index
        Case 1   '詢問進度函
           If CheckUse("frm050301", strExec) = True Then
              frm050301.Show
           End If
        Case 2   '期限管制表
           If CheckUse("frm050302", strExec) = True Then
              frm050302.Show
           End If
        Case 3   '代理人案件收達/提申管制表
           If CheckUse("frm050303", strExec) = True Then
              frm050303.Show
           End If
        Case 4   '收文未發文明細表
           If CheckUse("frm050304", strExec) = True Then
              frm050304.Show
           End If
        Case 5   '催審表
           If CheckUse("frm050305", strExec) = True Then
              frm050305.Show
           End If
        Case 6   '智權人員案件明細表
           If CheckUse("frm050306", strExec) = True Then
              frm050306.Show
           End If
        Case 7   '收文簿
           If CheckUse("frm050307", strExec) = True Then
              StrStartSystemByNick = "CFP"
              frm050307.Show
           End If
        Case 8   '收文明細表
           If CheckUse("frm050308", strExec) = True Then
              frm050308.Show
           End If
        Case 9   '承辦人發文明細表
           If CheckUse("frm050309", strExec) = True Then
              frm050309.Show
           End If
        Case 10  '發文點數明細表
           If CheckUse("frm050310", strExec) = True Then
              frm050310.Show
           End If
          'Add By Cheng 2003/02/05
        Case 11  '新案承辦人明細表
           If CheckUse("frm050321", strExec) = True Then
              frm050321.Show
           End If
        Case 12  '承辦人准駁明細表
           If CheckUse("frm050311", strExec) = True Then
              frm050311.Show
           End If
        Case 13  '期限通知管制表
           If CheckUse("frm050312", strExec) = True Then
              frm050312.Show
           End If
        Case 14  '取消收文明細表
           If CheckUse("frm050313", strExec) = True Then
              frm050313.Show
           End If
        Case 15  '後金案件表
           If CheckUse("frm050314", strExec) = True Then
              frm050314.Show
           End If
        Case 16  '延期明細表
           If CheckUse("frm050315", strExec) = True Then
              frm050315.Show
           End If
        Case 17  '代理人案件總簿
           If CheckUse("frm050316", strExec) = True Then
              frm050316.Show
           End If
        Case 18  '客戶案件總簿
           If CheckUse("frm050317", strExec) = True Then
              'Modify by Amy 2017/08/04 原:frm050317
              frm0503171.Show
           End If
        Case 19  '代理人/申請人名單
           If CheckUse("frm050318", strExec) = True Then
              frm050318.Show
           End If
'edit by nickc 2005/11/10
'        Case 20
'           If CheckUse("frm040320", strExec) = True Then
'              frm040320.Show
'           End If
        Case 21  '地址條列印
           If CheckUse("frm083014", strExec) = True Then
              frm083014.Show
           End If
        'Add By Cheng 2003/02/05
        Case 22  'TNT列印
            If CheckUse("frm060321", strExec) = True Then
                frm060321.Show
            End If
        'Add by Lydia 2014/12/26
        Case 23  'DHL列印
           If CheckUse("frm060330", strExec) = True Then
              frm060330.Show
           End If
        'Add by Morgan 2007/12/26
        Case 24  '美國發明退公開費報表/指示信
           If CheckUse("frm050325", strExec) = True Then
              frm050325.Show
           End If
        'Add by Sindy 2012/2/29
        Case 25  '未收文期限提醒E-Mail
           If CheckUse("frm050326", strExec) = True Then
              frm050326.Show
              frm050326.Text1(1) = "CFP"   'add by sonia 2016/1/8
           End If
        'Added by Morgan 2018/10/25
        Case 26 '期限通知檢核及報表列印
            If CheckUse("frm040335CFP", strExec) = True Then
               frm040335.m_ProState = "CFP"
               frm040335.Show
            End If
        
    End Select
End Sub

Private Sub mnu0504_Click(Index As Integer)
   intPCaseKind = 1
   intPWhere = 1
   ToolHide
   Select Case Index
      Case 1   '承辦人收文統計表
         If CheckUse("frm050401", strExec) = True Then
            frm050401.Show
         End If
      Case 2   '承辦人發文統計表
         If CheckUse("frm050402", strExec) = True Then
            frm050402.Show
         End If
      Case 3   '准駁統計表
         If CheckUse("frm050403", strExec) = True Then
            frm050403.Show
         End If
      Case 4   '代理人新案案件統計表
         If CheckUse("frm050404", strExec) = True Then
            frm050404.Show
         End If
        'Add By Cheng 2003/12/03
      Case 5   '代理人案件年度統計表
         If CheckUse("frm050407", strExec) = True Then
            frm050407.Show
         End If
        'End
      Case 6   '代理人/申請人新申請案排行榜
         If CheckUse("frm050405", strExec) = True Then
            StrStartSystemByNick = "CFP"
            frm050405.Show
         End If
      Case 7   '逾期未結案統計表
         If CheckUse("frm084004", strExec) = True Then
            frm084004.Tag = 2
            frm084004.Show
         End If
      'add by sonia 2015/10/2
      Case 8 '互惠代理人案件統計表
         If CheckUse("frm050408", strExec) = True Then
            frm050408.Show
         End If
      Case Else
   End Select
End Sub

Private Sub mnu0505_Click(Index As Integer)
   ToolHide
   strSysKind = "1"
   Select Case Index
   Case 1   '專利案件基本資料維護
      If CheckUse("frm050701", strExec) = True Then
         strSysKind = "CFP"
         frm050701.Show
      End If
   Case 2   '服務業務基本資料維護
      If CheckUse("frm050702", strExec) = True Then
         strSysKind = "CPS"
         frm050702.Show
      End If
   Case 3   '案件進度檔資料維護
      If CheckUse("frm075004_1", strExec) = True Then
        'Add By Cheng 2002/11/04
        strSysKind = "CFP"
         frm075004_1.Show
      End If
   Case 4   '下一程序資料
      If CheckUse("frm075007_1", strExec) = True Then
        'Add By Cheng 2002/11/04
        strSysKind = "CFP"
         frm075007_1.Show
      End If
   Case 5   '國外代理人資料
      If CheckUse("frm050705", strExec) = True Then
         strSysKind = "CFP"
         frm050705.Show
      End If
   Case 6   '變更事項
      If CheckUse("frm050706", strExec) = True Then
         strSysKind = "CFP"
         frm050706.Show
      End If
   Case 7   '延期記錄資料維護
      If CheckUse("frm050707", strExec) = True Then
         strSysKind = "CFP"
         frm050707.Show
      End If
   Case 8   '案件國家收費表維護
      If CheckUse("frm12040102", strExec) = True Then
         strSysKind = "CFP"
         frm12040102.Show
      End If
   Case 9   '客戶發明人資料維護
      If CheckUse("frm050709", strExec) = True Then
         strSysKind = "CFP"
         frm050709.Show
      End If
   Case 11  '代理人變更名稱作業
      If CheckUse("frm140103", strExec) = True Then
         frm140103.Show
      End If
   Case 12  '申請人國外ID資料維護
      If CheckUse("frm050711", strExec) = True Then
         frm050711.Show
      End If
   Case 13  '客戶基本資料維護
      If CheckUse("frm140401", strExec) = True Then
         frm140401.Show
      End If
    'add by nick 2004/07/14
    Case 14  '客戶減免身分維護
     If CheckUse("frm050715", strExec) = True Then
        frm050715.Show
     End If
    'add by nickc 2007/10/17
    Case 15  '特殊人員設定
     If CheckUse("frm050716", strExec) = True Then
        frm050716.Show
     End If
    Case 16  '國外代理人目標給案量維護
      If CheckUse("frm140406", strExec) = True Then
         frm140406.Show
      End If
    Case 17  'CFP領證報價資料維護
      If CheckUse("frm12040114", strExec) = True Then
         frm12040114.Show
      End If
   'Add By Sindy 2012/4/10
   Case 18 '非本所實質客戶資料維護
      If CheckUse("frm12040155", strExec) = True Then
         frm12040155.Show
      End If
   'Add By Sindy 2012/12/19
   Case 19 'CF代理人報價附件維護
      If CheckUse("frm050717", strExec) = True Then
         frm050717.Show
      End If
   'Added by Lydia 2016/10/24
   Case 20 '申請人指定國外代理人維護
      If CheckUse("frm050718", strExec) = True Then
         frm050718.Show
      End If
   'Added by Lydia 2017/06/06
   Case 21  'CFP核駁報價資料維護
      If CheckUse("frm12040117", strExec) = True Then
         frm12040117.Show
      End If
   'Added by Lydia 2025/01/21
   Case 22  'CFP維持費/延展費資料維護
      If CheckUse("frm12040164", strExec) = True Then
         frm12040164.Show
      End If
   End Select
End Sub

Private Sub mnu09_Click(Index As Integer)
    ToolHide
    Select Case Index
    Case 4   '撰寫信函
        frm090401.Show
    'Add by Morgan 2007/12/25
    Case 5   'P案國外指示信
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
    Case 6 'P案各式申請書
      If CheckUse("frm04010301_1", strExec) = True Then
         frm04010301_1.Show
      End If
    'Remove by Lydia 2023/09/11 內專系統優化：原「承辦人作業」變更為「程序」，移除下列功能表。
    'Case 7   '聯絡單列印及E-Mail  '2011/9/22 加入
     '    frm1106.Show
     'end 2023/09/11
   'Modified by Lydia 2015/11/05 改為主管機關處理記錄
'    'Added by Morgan 2012/4/17
'    Case 10 '主管機關來電處理記錄
'      If CheckUse("frm04010515", strExec) = True Then
'         frm04010515.Show
'      End If
    'end 2015/11/05
   'Added by Morgan 2014/4/28
   Case 11 '公文來函判發作業
      
      '考慮職代問題,不必鎖權限
      'Modified by Morgan 2015/4/21 改回用權限控制,目前為游經理及王副總可執行
      'Modified by Morgan 2016/6/17 配合非臺灣案有其他判發人再改成可看自己及當時請假之被代理人
      'If CheckUse("frm040113", strExec) = True Then
         frm040113.Show
      'End If
      
   'Added by Morgan 2014/12/17
   Case 12 '發後補看作業
      '考慮職代問題,不必鎖權限
      'Added by Lydia 2015/04/22 分成兩個作業
      'frm040117.Show
      
   'Added by Sindy 2015/1/23
   Case 13 '結案單審核作業
      'frm040118.Show 'Mark by Amy 2018/08/29 改至一般作業
      
   'Added by Morgan 2015/11/4
   Case 14 '指示信判發作業
      If CheckUse("frm040119", strExec) = True Then
         frm040119.Show
      End If
      
   'Added by Sindy 2016/9/12
   Case 15 '專利處收件夾信件處理
      If CheckUse("frm04010518", strExec) = True Then
         frm04010518.Show
      End If
   
   'Add By Sindy 2017/9/20
   Case 16 '郵件分信關鍵字對照表維護
      If CheckUse("frm06010614", strExec) = True Then
         frm06010614.m_strLK12 = "P"
         frm06010614.Show
      End If
   End Select
End Sub

'Added by Lydia 2015/04/22 發後補看作業
'Remove by Lydia 2023/09/11 內專系統優化：原「承辦人作業」變更為「程序」，移除下列功能表。
'Private Sub mnu0912_Click(Index As Integer)
'    Select Case Index
'       '考慮職代問題,不必鎖權限
'       Case 1 '函知客戶
'           frm040117.Show
'       Case 2 '內部收文
'           frm040117_1.Show
'       Case Else
'    End Select
'End Sub
'end 2023/09/11

'Remove by Lydia 2023/09/11 內專系統優化：原「承辦人作業」變更為「程序」，移除下列功能表。
''Remove by Morgan 2010/12/29
''因畫面不夠用，本系統取消商標委查作業
'Private Sub mnu0902_Click(Index As Integer)
'Dim bolNoCheck As Boolean 'Added by Morgan 2013/10/8
'Dim nFrm As Form 'Add By Sindy 2018/1/24
'
''ProState = "1"
''ProSysState = "1"
'ToolHide
'Select Case Index
'   Case 1   '工作進度資料維護
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
'         End If
'      Next
'      '2018/1/24 END
'
'      If CheckUse("frm090201_4", strExec) Then
'         ProState = "1"
'         ProSysState = "1"
'
'         'Added by Morgan 2013/10/8
'         If Left(Pub_StrUserSt03, 2) = "P1" Then
'            '第2次以上可選擇
'            strSql = "select * from executelog where el01='frm090201_a' and el02='" & strUserNum & "' and el03=" & strSrvDate(1)
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'            If intI = 1 Then
'               'Modified by Morgan 2016/4/19
'               'If MsgBox("是否執行 當天本所期限案件... 等功能", vbYesNo + vbQuestion + vbDefaultButton1, "功能！") = vbNo Then
'               If MsgBox("是否執行 3個工作天內達本所期限案件... 等功能", vbYesNo + vbQuestion + vbDefaultButton1, "功能！") = vbNo Then
'               'end 2016/4/19
'                  bolNoCheck = True
'               End If
'            End If
'         End If
'
'         If bolNoCheck = True Then
'            frm090201_2.Show
'         Else
'         'end 2013/10/8
'
'            '2009/11/12 modify by sonia 改寫法無資料不顯示畫面
'            'frm090201_4.Show
'            frm090201_4.StrMenu1   '當天本所期限案件資料,無資料時由frm090201_4的nextstep執行下一畫面
'            If frm090201_4.TextOk = True Then frm090201_4.Show
'            '2009/11/12 end
'
'         End If 'Added by Morgan 2013/10/8
'      End If
'   'Add By Sindy 2013/9/3
'   Case 2 '待核判區
'      'If CheckUse("frm090202_1", strExec) Then
'         frm090202_1.m_ProSysState = "1" '承辦人
'         frm090202_1.Show
'      'End If
'   'Add By Cheng 2003/06/17
'   Case 3   '承辦人支援記錄維護
'      If CheckUse("frm090623P", strExec) Then
'        ProState = "1"
'        ProSysState = "1"
'         frm090623.Show
'      End If
'    'Add By Morgan 2003/12/25
'   Case 4   '承辦人外出記錄維護
'      If CheckUse("frm090626", strExec) Then
'        ProState = "1"
'        ProSysState = "1"
'         frm090626.Show
'      End If
'    'Add By Cheng 2003/12/16
'   Case 5   '承辦人特殊案件記錄維護
'      If CheckUse("frm090627P", strExec) Then
'        ProState = "1"
'        ProSysState = "1"
'         frm090627.Show
'      End If
'
'   'Add by Morgan 2011/7/27
'   Case 6   '承辦人修改記錄維護
'      '目前控制只有中所能用
'      If pub_strUserOffice = "2" Or Pub_StrUserSt03 = "M51" Then
'         If CheckUse("frm090633P", strExec) Then
'            ProState = "1"
'            ProSysState = "1"
'            frm090633.Show
'         End If
'      End If
'
'   'Add by Morgan 2011/8/1
'   Case 7   '承辦人衍生記錄維護
'      '目前控制只有中所能用
'      If pub_strUserOffice = "2" Or Pub_StrUserSt03 = "M51" Then
'         If CheckUse("frm090634P", strExec) Then
'            ProState = "1"
'            ProSysState = "1"
'            frm090634.Show
'         End If
'      End If
'
'   Case 8   '未齊備,未完稿,未發文查詢
'      'frm090202_1.Show  CheckUse時於FormName後面加 1,2 區分個人及管理
'      If CheckUse("frm0906121", strExec) Then
'         ProState = "1"
'         ProSysState = "1"
'         frm090612.Show
'      End If
'   Case 9   '工作進度資料查詢
'      If CheckUse("frm090203_1", strExec) Then
'         ProState = "1"
'         ProSysState = "1"
'         frm090203_1.Show
'      End If
'   'Add By Cheng 2002/08/27 CheckUse時於FormName後面加 1,2 區分個人及管理
'    'Modify By Cheng 2003/07/30
''   Case 5 '承辦人目標資料查詢
''      If CheckUse("frm0906221", strExec) Then
''         frm090622.Show
''      End If
'   Case 10   '承辦人達成情形查詢
'      If CheckUse("frm0906081", strExec) Then
'         ProState = "1"
'         ProSysState = "1"
'         frm090608.Show
'      End If
'
''Removed by Morgan 2022/1/17 沒在用,刪除 (原來就沒顯示)
''   Case 11   '同仁評分作業
''      If CheckUse("frm090204_1", strExec) Then
''         ProState = "1"
''         ProSysState = "1"
''         frm090204_1.Show
''      End If
'
'   Case 12   ' 工作進度資料列印
'      If CheckUse("frm090205_1", strExec) Then
'         ProState = "1"
'         ProSysState = "1"
'         frm090205_1.Show
'      End If
'   'add by nickc 2005/03/27
'   Case 16  '專利處每週速度考核表
'      If CheckUse("frm090624P", strExec) Then
'         ProState = "1"
'         ProSysState = "1"
'         frm090624.Show
'      End If
'   Case 17  '月考核
'      If CheckUse("frm090616P", strExec) Then
'         ProState = "1"
'         ProSysState = "1"
'         frm090616_0.Show
'      End If
'   Case 18  '季考核
'      If CheckUse("frm090618P", strExec) Then
'         ProState = "1"
'         ProSysState = "1"
'         frm090618.Show
'      End If
'   Case Else
'End Select
'End Sub
'end 2023/09/11

'Remove by Lydia 2023/09/11 內專系統優化：原「承辦人作業」變更為「程序」，移除下列功能表。
'Private Sub mnu090205_Click(Index As Integer)
''ProState = "1"
''ProSysState = "2"
'ToolHide
'Select Case Index
'   '專利案例個人輸入作業
'   Case 1
'      If CheckUse("frm090206_1", strExec) Then
'        ProState = "1"
'        ProSysState = "2"
'         frm090206_1.Show
'      End If
'   '專利案例資料查詢
'   Case 2
'      If CheckUse("frm090207_1", strExec) Then
'        ProState = "1"
'        ProSysState = "2"
'         frm090207_1.Show
'      End If
'   '專利案例資料彙整
'   Case 3
'      If CheckUse("frm090217_1", strExec) Then
'        ProState = "1"
'        ProSysState = "2"
'         frm090217_1.Show
'      End If
'   '專利案例資料維護
'   Case 4
'      If CheckUse("frm090206_2", strExec) Then
'        ProState = "1"
'        ProSysState = "2"
'         frm090206_2.Show
'      End If
'   Case Else
'End Select
'End Sub
'end 2023/09/11

'Remove by Lydia 2023/09/11 內專系統優化：原「承辦人作業」變更為「程序」，移除下列功能表。
'Private Sub mnu090207_Click(Index As Integer)
''ProState = "1"
''ProSysState = "2"
'ToolHide
'Select Case Index
'   Case 1   '公報簡訊個人輸入作業
'      If CheckUse("frm090208_1", strExec) Then
'         ProState = "1"
'         ProSysState = "2"
'         frm090208_1.Show
'      End If
'   Case 2   '公報簡訊查詢列印
'      If CheckUse("frm090212_1", strExec) Then
'         ProState = "1"
'         ProSysState = "2"
'         frm090212_1.Show
'      End If
'   Case 3   '公報簡訊資料彙整作業
'      If CheckUse("frm090209_1", strExec) Then
'         ProState = "1"
'         ProSysState = "2"
'         frm090209_1.Show
'      End If
'   Case 4   '公報簡訊資料維護
'      If CheckUse("frm090210_1", strExec) Then
'         ProState = "1"
'         ProSysState = "2"
'         frm090210_1.Show
'      End If
'   Case 5   '公報簡訊索引資料維護
'      If CheckUse("frm090211_1", strExec) Then
'         ProState = "1"
'         ProSysState = "2"
'         frm090211_1.Show
'      End If
'End Select
'
'End Sub
'end 2023/09/11

'Remove by Lydia 2023/09/11 內專系統優化：原「承辦人作業」變更為「程序」，移除下列功能表。
'Private Sub mnu090208_Click(Index As Integer)
''ProState = "1"
''ProSysState = "2"
'ToolHide
'Select Case Index
'   Case 2   '期刊資料維護
'      If CheckUse("frm090213", strExec) Then
'         ProState = "1"
'         ProSysState = "2"
'         frm090213.Show
'      End If
'   Case 3   '期刊索引資料維護
'      If CheckUse("frm090214", strExec) Then
'         ProState = "1"
'         ProSysState = "2"
'         frm090214.Show
'      End If
'   Case 1   '期刊資料查詢列印
'      If CheckUse("frm090215_1", strExec) Then
'         ProState = "1"
'         ProSysState = "2"
'         frm090215_1.Show
'      End If
'   Case Else
'End Select
'
'End Sub

'Remove by Lydia 2023/09/11 內專系統優化：原「承辦人作業」變更為「程序」，移除下列功能表。
'Private Sub mnu0903_Click(Index As Integer)
'Dim bolNoCheck As Boolean 'Added by Morgan 2016/4/19
''ProState = "1"
''ProSysState = "2"
'ToolHide
'Select Case Index
'   Case 1   '個人工作進度資料維護－當天及前一工作天分案案件資料
'        'Modify By Cheng 2003/06/27
''      If CheckUse("frm090711", strExec) Then
''         frm090711.Show
''      End If
'      If CheckUse("frm090711_2", strExec) Then
'         ProState = "1"
'         ProSysState = "2"
'         'Modified by Morgan 2016/4/19 比照工程師彈3個工作天內達所限案件
'        'frm090711_2.Show
'        '第2次以上可選擇
'         strSql = "select * from executelog where el01='frm090201_4' and el02='" & strUserNum & "' and el03=" & strSrvDate(1)
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'         If intI = 1 Then
'            If MsgBox("是否執行 3個工作天內達本所期限案件功能", vbYesNo + vbQuestion + vbDefaultButton1, "功能！") = vbNo Then
'               bolNoCheck = True
'            End If
'         End If
'         If bolNoCheck = True Then
'            frm090711_2.Show
'         Else
'            frm090201_4.m_bolIsDrawer = True
'            frm090201_4.StrMenu1
'            If frm090201_4.TextOk = True Then frm090201_4.Show
'         End If
'        'end 2016/4/19
'      End If
'   'Add By Sindy 2013/9/3
'   Case 2 '待核判區
'      'If CheckUse("frm090202_1", strExec) Then
'         frm090202_1.m_ProSysState = "2" '繪圖人員
'         frm090202_1.Show
'      'End If
'   'Add By Cheng 2003/06/17
'   Case 3   '繪圖人員支援記錄維護
'      If CheckUse("frm090623P", strExec) Then
'         ProState = "1"
'         ProSysState = "2"
'         frm090623.Show
'      End If
'    'Add By Morgan 2003/12/16
'   Case 4   '繪圖人員外出記錄維護
'      If CheckUse("frm090626", strExec) Then
'         ProState = "1"
'         ProSysState = "2"
'         frm090626.Show
'      End If
'   Case 5   '未齊備、未完稿、未發文查詢
'      If CheckUse("frm090705", strExec) Then
'         ProState = "1"
'         ProSysState = "2"
'         frm090705.Show
'      End If
'   Case 6   '工作進度資料查詢
'      If CheckUse("frm090303_1", strExec) Then
'         ProState = "1"
'         ProSysState = "2"
'         frm090303_1.Show
'      End If
'
''Removed by Morgan 2022/1/17 沒在用,刪除 --翔龍、游經理確認
''   Case 7   '同仁評分作業
''      If CheckUse("frm090204_1", strExec) Then
''         ProState = "1"
''         ProSysState = "2"
''         frm090204_1.Show
''      End If
'
'   Case 8   '繪圖人員工作進度資料查詢
'      If CheckUse("frm090706", strExec) Then
'         ProState = "1"
'         ProSysState = "2"
'         frm090706.Show
'      End If
'    'Add By Cheng 2003/07/30
'   Case 9   '繪圖人員達成情形查詢
'      If CheckUse("frm0907011", strExec) Then
'         ProState = "1"
'         ProSysState = "2"
'         frm090701.Show
'      End If
'   'add by nickc 2005/03/27
'   Case 10   '專利處每週速度考核表
'      If CheckUse("frm090624P", strExec) Then
'         ProState = "1"
'         ProSysState = "2"
'         frm090624.Show
'      End If
'   Case 11  '月考核
'      If CheckUse("frm090616D", strExec) Then
'         ProState = "1"
'         ProSysState = "2"
'         frm090616_0.Show
'      End If
'   Case 12  '季考核
'      If CheckUse("frm090618D", strExec) Then
'         ProState = "1"
'         ProSysState = "2"
'         frm090618.Show
'      End If
'End Select
'End Sub
'end 2023/09/11

Private Sub mnu0907_Click(Index As Integer)
   ToolHide
   Select Case Index
'      'Add By Sindy 2013/5/16
'      Case 6 '待核判區
'         If CheckUse("frm090202_1", strExec) Then
'            frm090202_1.m_ProSysState = "1" '承辦人
'            frm090202_1.Show
'         End If
   End Select
'''''edit by nickc 2007/12/12 專利處修改
'''''''ProState = "2"
'''''''ProSysState = "1"
''''''ToolHide
''''''Select Case Index
''''''Case 2
''''''   If CheckUse("frm090606", strExec) Then
''''''    ProState = "2"
''''''    ProSysState = "1"
''''''      frm090606.Show
''''''   End If
''''''Case Else
''''''
''''''End Select
End Sub

'Remove by Lydia 2023/09/11 內專系統優化：原「承辦人作業」變更為「程序」，移除下列功能表。
'Private Sub mnu090703_Click(Index As Integer)
''''''edit by nickc 2007/12/12 專利處修改
''...原程式碼已清除
'
'ToolHide
'Select Case Index
'   Case 1   '承辦人工作進度資料查詢
'      If CheckUse("frm090614", strExec) Then
'       ProState = "2"
'       ProSysState = "1"
'         frm090614.Show
'      End If
'   Case 2   '承辦人達成情形查詢
'      If CheckUse("frm0906082", strExec) Then
'       ProState = "2"
'       ProSysState = "1"
'         frm090608.Show
'      End If
'   Case 3   '承辦人工作量查詢
'      If CheckUse("frm090609", strExec) Then
'       ProState = "2"
'       ProSysState = "1"
'         frm090609.Show
'      End If
'   Case 4   '承辦人每日分案情形查詢
'      If CheckUse("frm090610", strExec) Then
'       ProState = "2"
'       ProSysState = "1"
'         frm090610.Show
'      End If
'   Case 5   '承辦天數統計查詢
'      If CheckUse("frm090611", strExec) Then
'       ProState = "2"
'       ProSysState = "1"
'         frm090611.Show
'      End If
'   Case 6   '未齊備,未完稿,未發文查詢 'CheckUse時於FormName後面加 1,2 區分個人及管理
'      If CheckUse("frm0906122", strExec) Then
'       ProState = "2"
'       ProSysState = "1"
'         frm090612.Show
'      End If
'   Case 7   '案件處理時間統計查詢
'      If CheckUse("frm090613", strExec) Then
'       ProState = "2"
'       ProSysState = "1"
'         frm090613.Show
'      End If
'   Case 8   '工程師每週完稿明細
'      If CheckUse("frm090625", strExec) Then
'       ProState = "2"
'       ProSysState = "1"
'         frm090625.Show
'      End If
'   Case 9   '案件逾期及異常查詢 (已發文未輸會稿完成日查詢)
'      If CheckUse("frm090628", strExec) Then
'       ProState = "2"
'       ProSysState = "1"
'         frm090628.Show
'      End If
'   Case 10  '加乘註記修改歷史查詢列印
'      If CheckUse("frm090630", strExec) Then
'       ProState = "2"
'       ProSysState = "1"
'         frm090630.Show
'      End If
'   Case 11  '英文核稿查詢
'         If CheckUse("frm090218", strExec) Then
'           ProState = "2"
'           ProSysState = "1"
'            frm090218.Show
'         End If
'   Case 12  '智權人員收文高低標查詢
'      If CheckUse("frm090607", strExec) Then
'       ProState = "2"
'       ProSysState = "1"
'         frm090607.Show
'      End If
'   'Add by Morgan 2010/10/12
'   Case 13  '預定會稿日異常案件查詢
'      If CheckUse("frm090632", strExec) Then
'         ProState = "2"
'         ProSysState = "1"
'         frm090632.Show
'      End If
'   'Added by Lydia 2016/01/25
'   Case 14   '支援記錄獎金統計
'      If CheckUse("frm090639", strExec) Then
'         frm090639.Show
'      End If
'   'Add by Amy 2017/12/20
'  Case 15 '待辦案件量統計查詢
'   If CheckUse("frm090641", strExec) Then
'      frm090641.Show
'   End If
'   Case Else
'End Select
'End Sub
'end 2023/09/11

'Remove by Lydia 2023/09/11 內專系統優化：原「承辦人作業」變更為「程序」，移除下列功能表。
'Private Sub mnu090704_Click(Index As Integer)
'''''edit by nickc 2007/12/12 專利處修改
'...原程式碼已清除
'ToolHide
'Select Case Index
'   Case 1   '專利處每週速度考核表
'      If CheckUse("frm090624", strExec) Then '專利處每週速度考核表
'         ProState = "2"
'         ProSysState = "1"
'         frm090624.Show
'      End If
'   Case 2   '月考核
'      If CheckUse("frm090616M", strExec) Then
'         ProState = "2"
'         ProSysState = "1"
'         frm090616_0.Show
'      End If
'   Case 3   '季考核
'      If CheckUse("frm090618M", strExec) Then
'         ProState = "2"
'         ProSysState = "1"
'         frm090618.Show
'      End If
'   Case 4   '工程師每月目標基數設定
'      If CheckUse("frm090615", strExec) Then
'         ProState = "2"
'         ProSysState = "1"
'         frm090631.Show
'      End If
'   Case 5   '個人、繪圖人員目標資料維護
'      If CheckUse("frm090615", strExec) Then
'         ProState = "2"
'         ProSysState = "1"
'         frm090615.Show
'      End If
'   Case 6   '獎金輸入
'      If CheckUse("frm090617", strExec) Then
'         ProState = "2"
'         ProSysState = "1"
'         frm090617.Show
'      End If
'   Case 7   '獎金明細表
'      If CheckUse("frm090619", strExec) Then
'         ProState = "2"
'         ProSysState = "1"
'         frm090619.Show
'      End If
'   Case Else
'End Select
'End Sub
'end 2023/09/11

'Remove by Lydia 2023/09/11 內專系統優化：原「承辦人作業」變更為「程序」，移除下列功能表。
'Private Sub mnu090705_Click(Index As Integer)
'ToolHide
'Select Case Index
'   Case 1   '承辦人支援記錄維護
'      If CheckUse("frm090623M", strExec) Then
'         ProState = "2"
'         ProSysState = "1"
'         frm090623.Show
'      End If
'   Case 2   '承辦人特殊案件記錄維護
'      If CheckUse("frm090627M", strExec) Then
'         ProState = "2"
'         ProSysState = "1"
'         frm090627.Show
'      End If
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
'   Case 5  '國內外案件資料維護
'      If CheckUse("frm050106_1", strExec) = True Then
'         ProState = "2"
'         ProSysState = "1"
'         frm050106_1.intWhereToGo = 0
'         frm050106_1.Show
'      End If
'   Case 6   '每月目次重編作業
'       If CheckUse("frm090606", strExec) Then
'         ProState = "2"
'         ProSysState = "1"
'         frm090606.Show
'      End If
'
''Memo by Morgan 2018/5/23 操作超慢,目前沒用選單已設不顯示
''Removed by Morgan 2022/1/17 刪除
''   Case 7   '承辦人,核稿人對照資料維護
''      If CheckUse("frm090621", strExec) Then
''         ProState = "2"
''         ProSysState = "1"
''         frm090621.Show
''      End If
''end 2022/1/7
'
'   Case 8   '特殊加乘註記維護
'      If CheckUse("frm090629", strExec) Then
'         ProState = "2"
'         ProSysState = "1"
'         frm090629.Show
'      End If
'   '英文核稿人欄修改權限設定
'   Case 9
'      If CheckUse("frm090202_6", strExec) Then
'         ProState = "2"
'         ProSysState = "1"
'         frm090202_6.Show
'      End If
'   'Added by Lydia 2016/12/29
'   Case 10  '免費修正事由維護
'      If CheckUse("frm090640", strExec) Then
'         ProState = "2"
'         ProSysState = "1"
'         frm090640.Show
'      End If
'   Case Else
'End Select
'End Sub
'end 2023/09/11

'Remove by Lydia 2023/09/11 內專系統優化：原「承辦人作業」變更為「程序」，移除下列功能表。
'Private Sub mnu0908_Click(Index As Integer)
''ProState = "2"
''ProSysState = "2"
'ToolHide
'Select Case Index
'   'Add By Cheng 2003/06/17
'   Case 3   '繪圖人員支援記錄維護
'   '    ProState = "3"
'   '    ProSysState = "2"
'       If CheckUse("frm090623M", strExec) Then
'         ProState = "3"
'         ProSysState = "2"
'         frm090623.Show
'       End If
'   Case 4   '繪圖分案作業
'   '    ProState = "3"
'   '    ProSysState = "2"
'       If CheckUse("frm090712", strExec) Then
'         ProState = "3"
'         ProSysState = "2"
'         frm090712.Show
'       End If
''   'Add By Sindy 2013/5/16
''   Case 5 '待核判區
''      If CheckUse("frm090202_1", strExec) Then
''         frm090202_1.m_ProSysState = "2" '繪圖人員
''         frm090202_1.Show
''      End If
'   Case Else
'End Select
'End Sub
'end 2023/09/11

'Modified by Morgan 2012/9/25 調整順序
'Remove by Lydia 2023/09/11 內專系統優化：原「承辦人作業」變更為「程序」，移除下列功能表。
'Private Sub mnu090801_Click(Index As Integer)
''ProState = "2"
''ProSysState = "2"
'ToolHide
'Select Case Index
'   Case 1   '繪圖人員工作進度資料查詢
'      'modify by sonia 2018/5/3 原未區分會因個人之工作進度資料列印權限而開放主管權限
'      'If CheckUse("frm090706", strExec) Then
'      If CheckUse("frm0907062", strExec) Then
'         ProState = "2"
'         ProSysState = "2"
'         frm090706.Show
'      End If
'
'   'Added by Morgan 2012/9/25
'   Case 2   '繪圖超時案件查詢
'      If CheckUse("frm090707", strExec) Then
'         ProState = "2"
'         ProSysState = "2"
'         frm090707.Show
'      End If
'   Case 3   '繪圖人員達成情形
'       'Modify By Cheng 2003/07/30
'       '表單名稱後加一碼以區別個人或管理權限
'   '   If CheckUse("frm090701", strExec) Then
'      If CheckUse("frm0907012", strExec) Then
'         ProState = "2"
'         ProSysState = "2"
'         frm090701.Show
'      End If
'   Case 4   '繪圖人員工作量查詢
'      If CheckUse("frm090702", strExec) Then
'         ProState = "2"
'         ProSysState = "2"
'         frm090702.Show
'      End If
'   Case 5   '繪圖人員每日分案情形查詢
'      If CheckUse("frm090703", strExec) Then
'         ProState = "2"
'         ProSysState = "2"
'         frm090703.Show
'      End If
'   Case 6   '繪圖人員作業天數統計查詢
'      If CheckUse("frm090704", strExec) Then
'         ProState = "2"
'         ProSysState = "2"
'         frm090704.Show
'      End If
'   Case 7   '未齊備、未完稿、未發文查詢
'      If CheckUse("frm090705", strExec) Then
'         ProState = "2"
'         ProSysState = "2"
'         frm090705.Show
'      End If
'   Case Else
'End Select
'End Sub
'end 2023/09/11

'Remove by Lydia 2023/09/11 內專系統優化：原「承辦人作業」變更為「程序」，移除下列功能表。
'Private Sub mnu090802_Click(Index As Integer)
''ProState = "2"
''ProSysState = "2"
'ToolHide
'Select Case Index
'   Case 1   '個人、繪圖人員目標資料維護
'      If CheckUse("frm090615", strExec) Then
'         ProState = "2"
'         ProSysState = "2"
'         frm090615.Show
'      End If
'   Case 2   '月考核
'      'edit by nickc 2005/03/07
'      'If CheckUse("frm090709", strExec) Then
'      If CheckUse("frm090616DM", strExec) Then
'         ProState = "2"
'         ProSysState = "2"
'         'edit by nickc 2005/03/07
'         'frm090709.Show
'         frm090616_0.Show
'      End If
'   Case 3   '季考核
'      'edit by nickc 2005/03/07
'      'If CheckUse("frm090710", strExec) Then
'      If CheckUse("frm090618DM", strExec) Then
'         ProState = "2"
'         ProSysState = "2"
'         'edit by nickc 2005/03/07
'         'frm090710.Show
'         frm090618.Show
'      End If
'   Case Else
'End Select
'End Sub
'end 2023/09/11

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

Private Sub mnu11_Click(Index As Integer)
   ToolHide
   Select Case Index
      'Add By Cheng 2002/12/09
      Case 3   '來文資料稽核表
         If CheckUse("frm12040121", strExec) = True Then
            frm12040121.Show
         End If
      'Add By Cheng 2002/08/20
      Case 4   '聯絡單列印及E-Mail
         frm1106.Show
      Case 5   '相關卷號
         If CheckUse("frm1103_1", strExec) = True Then
            frm1103_1.Show
         End If
      Case 6   '多國案卷號關係建立
         If CheckUse("frm1104", strExec) = True Then
            frm1104.Show
         End If
      'Add By Cheng 2004/03/16
      Case 7   '分割案件關係維護
         If CheckUse("frm02010604_1", strExec) = True Then
            frm02010604_1.Show
         End If
      'Add By Cheng 2003/06/26
      Case 8   '撰寫信函作業
'         If CheckUse("frm090401", strExec) = True Then
            frm090401.Show
'         End If
      'Add by Morgan 2004/10/13
      Case 14  '部門別送件清單列印
         If CheckUse("frm1108", strExec) = True Then
            frm1108.Show
         End If
      
      'Add by Morgan 2011/6/2
      Case 15  '部門別電子送件清單列印
         If CheckUse("frm1109", strExec) = True Then
            frm1109.Show
         End If
         
      'add by nickc 2005/05/03
      Case 16  '銷案延遲日期輸入作業
         If CheckUse("frm140501", strExec) = True Then
            frm140501.Show
         End If
      'add by nickc 2005/07/22 CF 結餘單查詢
      Case 17
         If CheckUse("frm040202", strExec) = True Then
            frm040202.Show
         End If
      'add by nickc 2005/07/22
      Case 18   'CF 結餘資料維護
         If CheckUse("frm040206", strExec) = True Then
            frm040206.Show
         End If
      'add by nickc 2008/03/27
      Case 19   'CF 結餘單案件明細查詢
         If CheckUse("frm040208", strExec) = True Then
            frm040208.Show
         End If
      '2011/3/29 ADD BY SONIA
'      Case 20   '變更FC代理人作業
'         If CheckUse("frm110104_1", strExec) = True Then
'            frm110104_1.Show
'         End If
'move by Lydia 2015/10/14 從Law移到Patpro之共同程序
      '客戶應收帳款收文檢查上限
      Case 21 'Memo by Lydia 2020/02/07 隱藏: 改由財務室輸入
         If CheckUse("frm140502", strExec) = True Then
            frm140502.Show
         End If
      '客戶預定收款日放寬月數上限
      Case 22
         If CheckUse("frm140503", strExec) = True Then
            frm140503.Show
         End If
'end 2015/10/14
      'Added by Lydia 2018/08/17
      Case 23    '客戶特殊付款週期維護  'Memo by Lydia 2020/02/07 隱藏: 改由財務室輸入
         If CheckUse("frm140504", strExec) = True Then
            frm140504.Show
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
 'Add by Lydia 2014/10/14 FMP案(補正203和第一階段請求110)可不必輸入申請案號和證書,預設不印指示信和不詢問是否要作結餘
      Case 4   'ＦＭＰ解除期限 = 內專程序特定解除期限
         If CheckUse("frm110101_3", strExec) = True Then
            frm110101_3.Show
         End If
   End Select
End Sub

Public Sub mnu1102_Click(Index As Integer)
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
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   ToolShow
   tool1_enabled
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
      'Add by Morgan 2010/11/19
      'Memo by Morgan 2025/6/11 目前沒開放給User操作(選單已設不顯示)
      Case 13  '客製化請款項目資料維護
         If CheckUse("Frmacc21t0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Frmacc21t0.Show
      'Add by Morgan 2016/4/12
      Case 14  '帳單輸入-整批
         If CheckUse("Frmacc21u0", strExec) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Me.MousePointer = vbHourglass
         tool3_enabled
         Frmacc21u0.Show
         Me.MousePointer = vbDefault
   End Select
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
         '2007/11/29 ADD BY SONIA 外專使用者預設條件
         If Mid(GetStaffDepartment(strUserNum), 1, 2) = "F2" Then
            Frmacc24i0.Text7 = "F20"
            Frmacc24i0.Text8 = "F29"
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
         '2007/11/6 ADD BY SONIA 外專使用者預設條件
         If Mid(GetStaffDepartment(strUserNum), 1, 2) = "F2" Then
            Frmacc24c0.Text9 = "F20"
            Frmacc24c0.Text10 = "F29"
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
      'Add by Morgan 2010/8/24
      Case 9   '國外請款點數分析表
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
       Case 1   '作業失誤資料維護
          If CheckUse("frm050714", strExec) = True Then
             frm050714.Show
          End If
       Case 2   '失業失誤明細表
          If CheckUse("frm040327", strExec) = True Then
             frm040327.Show
          End If
    End Select
End Sub

Private Sub mnu15_Click(Index As Integer)
   ToolHide
   Select Case Index
      'Add by Morgan 2010/2/3
      Case 0   '印表機設定
         frm880011.bolAppOnly = True
         frm880011.Show 1
         
      'Add by Morgan 2008/3/27
      Case 1   '報表紙張格式設定
         frm880013.Show vbModal
      
      'Added by Morgan 2015/3/19
      Case 2 '解除畫面擷取限制
         frmChgUser.Caption = "解除畫面擷取限制"
         frmChgUser.SSTab1.TabVisible(1) = True
         frmChgUser.SSTab1.TabVisible(0) = False
         frmChgUser.Show
         
   End Select
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
      Case 2   '定稿資料維護
         If CheckUse("frm1105", strExec) = True Then
            frm1105.Show
         End If
   End Select
End Sub

Private Sub mnu1601_Click(Index As Integer)

   'Dim iDefaultPrinter As Integer 'Remove by Morgan 2010/2/3
   Dim strOsPrinter As String 'Added by Morgan 2025/9/8
   
   ToolHide
   
   'Add by Morgan 2006/10/19
   '設定控制台&Word預設印表機
   Load frm880011
   'Modify by Morgan 2010/2/3
   'iDefaultPrinter = frm880011.GetPrinterIndex
   'Modified by Morgan 2025/9/8 列印過程 pub_OsPrinter 會被改
   'pub_OsPrinter = PUB_GetOsDefaultPrinter
   strOsPrinter = PUB_GetOsDefaultPrinter
   'end 2025/9/8
   'end 2010/2/3
   frm880011.Show 1
   'end 2006/10/19
   
   Select Case Index
      Case 1 '橫式
          PrinterLetterDemand "1"
      Case 2 '英文
          PrinterLetterDemand "4"
      Case 3 '直式
          PrinterLetterDemand "2"
      Case 4 '日文
          PrinterLetterDemand "3"
      Case 5 '申請書
          PrinterLetterDemand "5"
      Case 6 '報價通知定稿
          'Modified by Morgan 2015/10/12 +m_UserNo
          PUB_Cache2Letter , , , , , , m_UserNo
      Case 7 '大陸指示信
          PrinterLetterDemand "6"
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
   'Modified by Morgan 2025/9/8 列印過程 pub_OsPrinter 會被改
   'PUB_SetOsDefaultPrinter pub_OsPrinter
   PUB_SetOsDefaultPrinter strOsPrinter
   'end 2025/9/8
   'end 2025/9/8
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
   Select Case Index
      'Modify By Cheng 2002/10/11
'      Case 16: PrinterLetterDemand
   End Select
End Sub

Private Sub Timer1_Timer()
'Add By Cheng 2002/11/13
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
    'add by nickc 20050818
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
    'add by nickc 2005/08/18
    'If mnu21(9).Enabled = False Then mnu21(9).Enabled = True
    
    'Modified by Morgan 2014/7/17
    'If mnu2102(2).Enabled = False Then mnu2102(2).Enabled = True
    If bXForm Then
      frmX.cmdOK(0).Enabled = True
      frmX.cmdOK(1).Enabled = True
    End If
    'end 2014/7/17
End If

'Add By Cheng 2002/11/19
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

Private Sub PrintData1()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strDate As String
Dim ii As Integer

If MsgBox("準備列印大陸案件齊備３天未完稿清單!!!", vbExclamation + vbOKCancel) = vbOK Then 'Added by Morgan 2016/9/21 改先問，否則不印也要等
    strDate = CompWorkDay(3, strSrvDate(1), 1)
'edit by nickc 2007/10/31 改成抓 table
'    StrSQLa = "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(PA57,Null,Null,'＊'), PA05, NA03, Decode(PA09,'020',CPM04,CPM03), CP48, CP06, EP06, ST02  From EngineerProgress, CaseProgress, Patent, Nation, CasePropertyMap, Staff Where EP02=CP09 And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 And PA09=NA01(+) And CP01=CPM01 And CP10=CPM02 And CP13=ST01(+) And CP27 Is Null And CP57 Is Null And PA57 Is Null And EP05='95014' And EP09 Is Null And EP06<" & strDate
'    StrSQLa = StrSQLa & " Union Select CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(SP15,Null,Null,'＊'), SP05, NA03, Decode(SP09,'020',CPM04,CPM03), CP48, CP06, EP06, ST02  From EngineerProgress, CaseProgress, ServicePractice, Nation, CasePropertyMap, Staff Where EP02=CP09 And CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04 And SP09=NA01(+) And CP01=CPM01 And CP10=CPM02 And CP13=ST01(+) And CP27 Is Null And CP57 Is Null And SP15 Is Null And EP05='95014' And EP09 Is Null And EP06<" & strDate
    'modify by sonia 2016/9/6 cp27,cp57改為cp158,cp159
    'Modified by Morgan 2016/9/22 調效能加 index IDXCP15815914 改語法
    'StrSQLa = "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(PA57,Null,Null,'＊'), PA05, NA03, Decode(PA09,'020',CPM04,CPM03), CP48, CP06, EP06, ST02  From EngineerProgress, CaseProgress, Patent, Nation, CasePropertyMap, Staff Where EP02=CP09 And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 And PA09=NA01(+) And CP01=CPM01 And CP10=CPM02 And CP13=ST01(+) And CP158=0 And CP159=0 And PA57 Is Null And instr('" & Pub_GetSpecMan("G") & "',ep05)>0 And EP09 Is Null And EP06<" & strDate
    'StrSQLa = StrSQLa & " Union Select CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(SP15,Null,Null,'＊'), SP05, NA03, Decode(SP09,'020',CPM04,CPM03), CP48, CP06, EP06, ST02  From EngineerProgress, CaseProgress, ServicePractice, Nation, CasePropertyMap, Staff Where EP02=CP09 And CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04 And SP09=NA01(+) And CP01=CPM01 And CP10=CPM02 And CP13=ST01(+) And CP158=0 And CP159=0 And SP15 Is Null And instr('" & Pub_GetSpecMan("G") & "',ep05)>0 And EP09 Is Null And EP06<" & strDate
    StrSQLa = "Select/*+ index(caseprogress idxcp15815914) */ CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(PA57,Null,Null,'＊'), PA05, NA03, Decode(PA09,'020',CPM04,CPM03), CP48, CP06, EP06, ST02  From EngineerProgress, CaseProgress, Patent, Nation, CasePropertyMap, Staff Where EP02(+)=CP09 And CP01=PA01(+) And CP02=PA02(+) And CP03=PA03(+) And CP04=PA04(+) and pa01 is not null And PA09=NA01(+) And CP01=CPM01(+) And CP10=CPM02(+) And CP13=ST01(+) And CP158=0 And CP159=0 And PA57 Is Null And instr('" & Pub_GetSpecMan("G") & "',cp14)>0 And EP09 Is Null And EP06<" & strDate
    StrSQLa = StrSQLa & " Union Select/*+ index(caseprogress idxcp15815914) */ CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(SP15,Null,Null,'＊'), SP05, NA03, Decode(SP09,'020',CPM04,CPM03), CP48, CP06, EP06, ST02  From EngineerProgress, CaseProgress, ServicePractice, Nation, CasePropertyMap, Staff Where EP02(+)=CP09 And CP01=SP01(+) And CP02=SP02(+) And CP03=SP03(+) And CP04=SP04(+) and sp01 is not null And SP09=NA01(+) And CP01=CPM01(+) And CP10=CPM02(+) And CP13=ST01(+) And CP158=0 And CP159=0 And SP15 Is Null And instr('" & Pub_GetSpecMan("G") & "',cp14)>0 And EP09 Is Null And EP06<" & strDate
    'end 2016/9/22
    StrSQLa = StrSQLa & " Order By 1 "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        'If MsgBox("準備列印大陸案件齊備３天未完稿清單!!!", vbExclamation + vbOKCancel) = vbOK Then 'Removed by Morgan 2016/9/21
            Page = 1
            PrintTitle1
            While Not rsA.EOF
                For ii = 0 To 7
                    strTemp1(ii) = "" & rsA.Fields(ii).Value
                Next ii
                PrintDatil1
                If iPrint > 10000 Then
                    Printer.CurrentX = PLeft1(0)
                    Printer.CurrentY = iPrint
                    Printer.Print String(200, "-")
                    rsA.MoveNext
                    If rsA.EOF = False Then
                        Printer.NewPage
                        Page = Page + 1
                        PrintTitle1
                    End If
                Else
                    rsA.MoveNext
                End If
            Wend
            Printer.EndDoc
    'Modified by Morgan 2016/9/21
    '    End If
    Else
         MsgBox "沒有大陸案件齊備３天未完稿案件!!!", vbInformation
    'end 2016/9/21
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
End If 'Added by Morgan 2016/9/21
End Sub

Private Sub PrintData2()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strDate As String
Dim ii As Integer
Dim strPromoter As String
        
        
If MsgBox("準備列印國外新申請案收文３天未齊備且無關聯案件清單!!!", vbExclamation + vbOKCancel) = vbOK Then
    strDate = CompWorkDay(3, strSrvDate(1), 1)
    '若為P的案件, 承辦人只抓韓聖文的資料91028->79075->94003
    'edit by nickc 2006/06/12
    'StrSQLa = "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(PA57,Null,Null,'＊') As F0, PA05 As F1, NA03 As F2, Decode(PA09,'020',CPM04,CPM03) As F3, CP05 As F4, CP06 As F5, ST02 As F6, CP14 As F7, CM01 As F8 From EngineerProgress, CaseProgress, Patent, Nation, Staff, Casepropertymap, CaseMap Where EP02=CP09 And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 And CP01=CM01(+) And CP02=CM02(+) And CP03=CM03(+) And CP04=CM04(+) And CP01=CPM01 And CP10=CPM02 And PA09=NA01 And CP13=ST01 And EP06 Is Null And CP01 In ('CFP', 'P') And CP31='Y' And CP27 Is Null And CP05>=20031201 And CP05<" & strDate & " And PA09>'000' And PA57 Is Null And '0'=CM10(+) And CP14=Decode(PA01,'P','79075',CP14) "
'edit by nickc 2007/10/31 改成抓table
'    StrSQLa = "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(PA57,Null,Null,'＊') As F0, PA05 As F1, NA03 As F2, Decode(PA09,'020',CPM04,CPM03) As F3, CP05 As F4, CP06 As F5, ST02 As F6, CP14 As F7, CM01 As F8 From EngineerProgress, CaseProgress, Patent, Nation, Staff, Casepropertymap, CaseMap Where EP02=CP09 And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 And CP01=CM01(+) And CP02=CM02(+) And CP03=CM03(+) And CP04=CM04(+) And CP01=CPM01 And CP10=CPM02 And PA09=NA01 And CP13=ST01 And EP06 Is Null And CP01 In ('CFP', 'P') And CP31='Y' And CP27 Is Null And CP05>=20031201 And CP05<" & strDate & " And PA09>'000' And PA57 Is Null And '0'=CM10(+) And CP14=Decode(PA01,'P','95014',CP14) and CP10 in (" & CaseMapOut & ") "
    'modify by sonia 2016/9/6 CP27 is null改為cp158=0
    'Modified by Morgan 2016/9/22 調效能加 index IDXCP15815914 改語法
    'StrSQLa = "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(PA57,Null,Null,'＊') As F0, PA05 As F1, NA03 As F2, Decode(PA09,'020',CPM04,CPM03) As F3, CP05 As F4, CP06 As F5, ST02 As F6, CP14 As F7, CM01 As F8 From EngineerProgress, CaseProgress, Patent, Nation, Staff, Casepropertymap, CaseMap Where EP02=CP09 And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 And CP01=CM01(+) And CP02=CM02(+) And CP03=CM03(+) And CP04=CM04(+) And CP01=CPM01 And CP10=CPM02 And PA09=NA01 And CP13=ST01 And EP06 Is Null And CP01 In ('CFP', 'P') And CP31='Y' And cp158=0 And CP05>=20031201 And CP05<" & strDate & " And PA09>'000' And PA57 Is Null And '0'=CM10(+) And instr(Decode(PA01,'P','" & Pub_GetSpecMan("G") & "',CP14),cp14)>0 and CP10 in (" & CaseMapOut & ") "
    'Modified by Morgan 2017/9/12 CFP剔除已設定無關聯P案者PA61='N'
    'Modify By Sindy 2023/3/31 拿掉,307 : 玫音說要將分割案排除（因分割案的關聯是透過與原案建關聯將案件串起來）。
    StrSQLa = "Select/*+ index(caseprogress idxcp15815905011014) */ CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(PA57,Null,Null,'＊') As F0, PA05 As F1, NA03 As F2, Decode(PA09,'020',CPM04,CPM03) As F3, CP05 As F4, CP06 As F5, ST02 As F6, CP14 As F7, CM01 As F8" & _
      " From EngineerProgress, CaseProgress, Patent, Nation, Staff, Casepropertymap, CaseMap Where EP02(+)=CP09 And CP01=PA01(+) And CP02=PA02(+) And CP03=PA03(+) And CP04=PA04(+)" & _
      " And CP01=CM01(+) And CP02=CM02(+) And CP03=CM03(+) And CP04=CM04(+) And CP01=CPM01 And CP10=CPM02 And PA09=NA01 And CP13=ST01 And EP06 Is Null" & _
      " And cp158=0 and cp159=0 And CP05>=20031201 And CP05<" & strDate & " And cp01='CFP' and CP10 in (101,102,103,104,105,109,110,112,113,114,115,118,122,201)" & _
      " And CP31='Y' And PA09>'000' And PA57 Is Null and pa61 is null And '0'=CM10(+)" & _
      " union all  Select/*+ index(caseprogress idxcp15815914) */ CP01||'-'||CP02||'-'||CP03||'-'||CP04||Decode(PA57,Null,Null,'＊') As F0, PA05 As F1, NA03 As F2, Decode(PA09,'020',CPM04,CPM03) As F3, CP05 As F4, CP06 As F5, ST02 As F6, CP14 As F7, CM01 As F8" & _
      " From EngineerProgress, CaseProgress, Patent, Nation, Staff, Casepropertymap, CaseMap Where EP02(+)=CP09 And CP01=PA01(+) And CP02=PA02(+) And CP03=PA03(+) And CP04=PA04(+)" & _
      " And CP01=CM01(+) And CP02=CM02(+) And CP03=CM03(+) And CP04=CM04(+) And CP01=CPM01 And CP10=CPM02 And PA09=NA01 And CP13=ST01 And EP06 Is Null" & _
      " And cp158=0 and cp159=0 and instr('" & Pub_GetSpecMan("G") & "',cp14)>0 And CP05+0>=20031201 And CP05+0<" & strDate & " And CP01||''='P' and CP10||'' in (101,102,103,104,105,109,110,112,113,114,115,118,122,201)" & _
      " And CP31='Y' And PA09>'000' And PA57 Is Null And '0'=CM10(+)"
    '2005/6/24 END
    StrSQLa = "Select * From ( " & StrSQLa & " ) T1 Where T1.F8 Is Null "
    'End
    StrSQLa = StrSQLa & " Order By 8, 1 "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
            'edit by nickc 2006/06/12
            'If MsgBox("準備列印國外新案收文３天未齊備且無關聯案件清單!!!", vbExclamation + vbOKCancel) = vbOK Then
            'If MsgBox("準備列印國外新申請案收文３天未齊備且無關聯案件清單!!!", vbExclamation + vbOKCancel) = vbOK Then 'Removed by Morgan 2016/9/21
                strPromoter = "" & rsA.Fields(7).Value
                Page = 1
                PrintTitle2 strPromoter
                While Not rsA.EOF
'                    '若CaseMap無資料
'                    If "" & rsA.Fields(8).Value = "" Then
                        If strPromoter <> "" & rsA.Fields(7).Value Then
                            Printer.CurrentX = PLeft1(0)
                            Printer.CurrentY = iPrint
                            Printer.Print String(200, "-")
'                            rsA.MoveNext
                            Printer.NewPage
                            Page = 1
                            strPromoter = "" & rsA.Fields(7).Value
                            PrintTitle2 strPromoter
                        End If
                        For ii = 0 To 6
                            strTemp1(ii) = "" & rsA.Fields(ii).Value
                        Next ii
                        PrintDatil2
                        If iPrint > 10000 Then
                            Printer.CurrentX = PLeft1(0)
                            Printer.CurrentY = iPrint
                            Printer.Print String(200, "-")
                            rsA.MoveNext
                            If rsA.EOF = False Then
                                Printer.NewPage
                                Page = Page + 1
                                PrintTitle2 strPromoter
                            End If
                        Else
                            rsA.MoveNext
                        End If
'                    '若CaseMap有資料
'                    Else
'                        rsA.MoveNext
'                    End If
                Wend
                Printer.EndDoc
            'End If 'Removed by Morgan 2016/9/21
'        End If
    'Added by Morgan 2016/9/21
    Else
      MsgBox "沒有國外新申請案收文３天未齊備且無關聯案件!!!", vbInformation
    'end 2016/9/21
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
End If 'Added by Morgan 2016/9/21
End Sub

Sub PrintTitle1()
   GetPleft1
   iPrint = 500
   Printer.Orientation = 2
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 18
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 4500
   Printer.CurrentY = iPrint
   Printer.Print "大陸案件齊備３天未完稿清單"
   iPrint = iPrint + 500
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.CurrentX = PLeft1(0)
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = 6000
   Printer.CurrentY = iPrint
   'modify by sonia 2016/9/6
   'Printer.Print "齊備日＜" & ChangeTStringToTDateString(ChangeWStringToTString(CompWorkDay(3, strSrvDate(1), 1) - 3))
   Printer.Print "齊備日＜" & ChangeTStringToTDateString(ChangeWStringToTString(CompWorkDay(3, strSrvDate(1), 1)))
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   iPrint = iPrint + 300
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "頁　　次：" & str(Page)
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft1(0)
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft1(0)
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft1(1)
   Printer.CurrentY = iPrint
   Printer.Print "案件名稱"
   Printer.CurrentX = PLeft1(2)
   Printer.CurrentY = iPrint
   Printer.Print "申請國家"
   Printer.CurrentX = PLeft1(3)
   Printer.CurrentY = iPrint
   Printer.Print "案件性質"
   Printer.CurrentX = PLeft1(4)
   Printer.CurrentY = iPrint
   Printer.Print "承辦期限"
   Printer.CurrentX = PLeft1(5)
   Printer.CurrentY = iPrint
   Printer.Print "本所期限"
   Printer.CurrentX = PLeft1(6)
   Printer.CurrentY = iPrint
   Printer.Print "齊備日"
   Printer.CurrentX = PLeft1(7)
   Printer.CurrentY = iPrint
   Printer.Print "智權人員"
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft1(0)
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
End Sub

Sub GetPleft1()
   Erase PLeft1
   PLeft1(0) = 500
   PLeft1(1) = PLeft1(0) + 2250
   PLeft1(2) = PLeft1(1) + 3875
   PLeft1(3) = PLeft1(2) + 1375
   PLeft1(4) = PLeft1(3) + 1625
   PLeft1(5) = PLeft1(4) + 1250
   PLeft1(6) = PLeft1(5) + 1250
   PLeft1(7) = PLeft1(6) + 1250
End Sub

Sub PrintDatil1()
    Printer.CurrentX = PLeft1(0)
    Printer.CurrentY = iPrint
    Printer.Print strTemp1(0)
    Printer.CurrentX = PLeft1(1)
    Printer.CurrentY = iPrint
    Printer.Print PUB_StrToStr(strTemp1(1), 30) '15個中文字
    Printer.CurrentX = PLeft1(2)
    Printer.CurrentY = iPrint
    Printer.Print PUB_StrToStr(strTemp1(2), 10) '5個中文字
    Printer.CurrentX = PLeft1(3)
    Printer.CurrentY = iPrint
    Printer.Print PUB_StrToStr(strTemp1(3), 12) '6個中文字
    Printer.CurrentX = PLeft1(4)
    Printer.CurrentY = iPrint
    Printer.Print ChangeTStringToTDateString(ChangeWStringToTString(strTemp1(4)))
    Printer.CurrentX = PLeft1(5)
    Printer.CurrentY = iPrint
    Printer.Print ChangeTStringToTDateString(ChangeWStringToTString(strTemp1(5)))
    Printer.CurrentX = PLeft1(6)
    Printer.CurrentY = iPrint
    Printer.Print ChangeTStringToTDateString(ChangeWStringToTString(strTemp1(6)))
    Printer.CurrentX = PLeft1(7)
    Printer.CurrentY = iPrint
    Printer.Print strTemp1(7)

    iPrint = iPrint + 300
End Sub

Sub GetPleft2()
   Erase PLeft1
   PLeft1(0) = 500
   PLeft1(1) = PLeft1(0) + 2250
   PLeft1(2) = PLeft1(1) + 3875
   PLeft1(3) = PLeft1(2) + 1375
   PLeft1(4) = PLeft1(3) + 1625
   PLeft1(5) = PLeft1(4) + 1250
   PLeft1(6) = PLeft1(5) + 1250
   PLeft1(7) = PLeft1(6) + 1250
End Sub

Sub PrintDatil2()
    Printer.CurrentX = PLeft1(0)
    Printer.CurrentY = iPrint
    Printer.Print strTemp1(0)
    Printer.CurrentX = PLeft1(1)
    Printer.CurrentY = iPrint
    Printer.Print PUB_StrToStr(strTemp1(1), 30) '15個中文字
    Printer.CurrentX = PLeft1(2)
    Printer.CurrentY = iPrint
    Printer.Print PUB_StrToStr(strTemp1(2), 10) '5個中文字
    Printer.CurrentX = PLeft1(3)
    Printer.CurrentY = iPrint
    Printer.Print PUB_StrToStr(strTemp1(3), 12) '6個中文字
    Printer.CurrentX = PLeft1(4)
    Printer.CurrentY = iPrint
    Printer.Print ChangeTStringToTDateString(ChangeWStringToTString(strTemp1(4)))
    Printer.CurrentX = PLeft1(5)
    Printer.CurrentY = iPrint
    Printer.Print ChangeTStringToTDateString(ChangeWStringToTString(strTemp1(5)))
    Printer.CurrentX = PLeft1(6)
    Printer.CurrentY = iPrint
    Printer.Print strTemp1(6)

    iPrint = iPrint + 300
End Sub

Sub PrintTitle2(ByVal strPromoter As String)
   GetPleft2
   iPrint = 500
   Printer.Orientation = 2
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 18
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 4500
   Printer.CurrentY = iPrint
   Printer.Print "國外新案收文３天未齊備且無關聯案件清單"
   iPrint = iPrint + 500
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.CurrentX = PLeft1(0)
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = 6000
   Printer.CurrentY = iPrint
   'modify by sonia 2016/9/6
   'Printer.Print "收文日＜" & ChangeTStringToTDateString(ChangeWStringToTString(CompWorkDay(3, strSrvDate(1), 1) - 3))
   Printer.Print "收文日＜" & ChangeTStringToTDateString(ChangeWStringToTString(CompWorkDay(3, strSrvDate(1), 1)))
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft1(0)
   Printer.CurrentY = iPrint
   Printer.Print "承辦人：" & GetStaffName(strPromoter)
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "頁　　次：" & str(Page)
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft1(0)
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft1(0)
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft1(1)
   Printer.CurrentY = iPrint
   Printer.Print "案件名稱"
   Printer.CurrentX = PLeft1(2)
   Printer.CurrentY = iPrint
   Printer.Print "申請國家"
   Printer.CurrentX = PLeft1(3)
   Printer.CurrentY = iPrint
   Printer.Print "案件性質"
   Printer.CurrentX = PLeft1(4)
   Printer.CurrentY = iPrint
   Printer.Print "收文日期"
   Printer.CurrentX = PLeft1(5)
   Printer.CurrentY = iPrint
   Printer.Print "本所期限"
   Printer.CurrentX = PLeft1(6)
   Printer.CurrentY = iPrint
   Printer.Print "智權人員"
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft1(0)
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
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

Public Sub ToolShow()
   Toolbar1.Visible = True
   StatusBar1.Visible = True
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
Private Sub mnu0910_Click(Index As Integer)
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

'Added by Lydia 2017/01/05 專利公報excel
Private Sub mnu0407_Click(Index As Integer)
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
   'Add by Amy 2024/01/22
   Case "frm090801_14" '接洽單-對造
      Set GetForm = frm090801_14
   Case "frm040104_1"
      Set GetForm = frm040104_1
   'Added by Morgan 2020/2/19
   Case "frm050709"
      Set GetForm = frm050709
   'Add By Sindy 2020/5/29
   Case "frm180301"
      Set GetForm = frm180301
   'Added by Morgan 2020/12/29
   Case "frm090401_1"
      Set GetForm = frm090401_1
   'Added by Sindy 2023/1/10
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
   Case "frm090201_2"
      Set GetForm = frm090201_2
   '2023/6/20 END
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
   'Add By Sindy 2024/6/28
   Case "frm090202_7" '申請人一個月內寄送資料
         Set GetForm = frm090202_7
   End Select
End Function
