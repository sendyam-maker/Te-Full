VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H80000018&
   Caption         =   "薪資系統"
   ClientHeight    =   4520
   ClientLeft      =   2630
   ClientTop       =   3560
   ClientWidth     =   9490
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  '最大化
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   1440
      Top             =   1410
   End
   Begin VB.Timer tmrConnect 
      Left            =   1485
      Top             =   2010
   End
   Begin VB.Timer Timer2 
      Left            =   270
      Top             =   1950
   End
   Begin VB.Timer Timer1 
      Left            =   270
      Top             =   1470
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '對齊表單下方
      Height          =   276
      Left            =   0
      TabIndex        =   1
      Top             =   4236
      Width           =   9492
      _ExtentX        =   16739
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
      Height          =   580
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9490
      _ExtentX        =   16739
      _ExtentY        =   1023
      ButtonWidth     =   614
      ButtonHeight    =   910
      Appearance      =   1
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
      Caption         =   "資料處理"
      Index           =   1
      Begin VB.Menu mnu1 
         Caption         =   "其他所得／扣款資料"
         Index           =   1
      End
      Begin VB.Menu mnu1 
         Caption         =   "同仁其他給付資料"
         Index           =   2
      End
      Begin VB.Menu mnu1 
         Caption         =   "每月獎金資料"
         Index           =   3
      End
      Begin VB.Menu mnu1 
         Caption         =   "薪資異動資料"
         Index           =   4
      End
      Begin VB.Menu mnu1 
         Caption         =   "員工借支資料"
         Index           =   5
      End
      Begin VB.Menu mnu1 
         Caption         =   "員工貸款資料"
         Index           =   6
      End
      Begin VB.Menu mnu1 
         Caption         =   "兼職人員每月工作時數資料"
         Index           =   7
      End
      Begin VB.Menu mnu1 
         Caption         =   "尾牙摸彩、年資、全勤獎金維護"
         Index           =   8
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "統計及批次作業"
      Index           =   2
      Begin VB.Menu mnu2 
         Caption         =   "每月薪資計算"
         Index           =   1
      End
      Begin VB.Menu mnu2 
         Caption         =   "婚喪互助扣款計算"
         Enabled         =   0   'False
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnu2 
         Caption         =   "勞、健保保費重算"
         Index           =   3
      End
      Begin VB.Menu mnu2 
         Caption         =   "銀行入帳媒體作業"
         Index           =   4
      End
      Begin VB.Menu mnu2 
         Caption         =   "補充保費繳納作業"
         Index           =   5
      End
      Begin VB.Menu mnu2 
         Caption         =   "補充保費明細申報檔"
         Index           =   6
      End
      Begin VB.Menu mnu2 
         Caption         =   "福利金轉檔"
         Index           =   7
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "查詢及列印"
      Index           =   3
      Begin VB.Menu mnu3 
         Caption         =   "查詢及檢核表"
         Index           =   1
         Begin VB.Menu mnu31 
            Caption         =   "每月薪資資料查詢"
            Index           =   1
         End
         Begin VB.Menu mnu31 
            Caption         =   "個人勞退自提明細表"
            Index           =   2
         End
         Begin VB.Menu mnu31 
            Caption         =   "薪資基本資料檢核表"
            Index           =   3
         End
         Begin VB.Menu mnu31 
            Caption         =   "人事薪資異動檢核表"
            Index           =   4
         End
         Begin VB.Menu mnu31 
            Caption         =   "員工應稅薪資檢核表"
            Index           =   5
         End
         Begin VB.Menu mnu31 
            Caption         =   "其他各類所得資料檢核表"
            Index           =   6
         End
         Begin VB.Menu mnu31 
            Caption         =   "股利及退職所得資料檢核表"
            Index           =   7
         End
      End
      Begin VB.Menu mnu3 
         Caption         =   "月報"
         Index           =   2
         Begin VB.Menu mnu32 
            Caption         =   "員工薪資明細表"
            Index           =   1
         End
         Begin VB.Menu mnu32 
            Caption         =   "員工薪資單"
            Index           =   2
         End
         Begin VB.Menu mnu32 
            Caption         =   "薪資所得稅明細表"
            Index           =   3
         End
         Begin VB.Menu mnu32 
            Caption         =   "薪資媒體入帳明細"
            Index           =   4
         End
         Begin VB.Menu mnu32 
            Caption         =   "薪資媒體轉帳遞送單及員工異動清冊"
            Index           =   5
         End
         Begin VB.Menu mnu32 
            Caption         =   "員工勞健保費及勞退自提明細表"
            Index           =   6
         End
         Begin VB.Menu mnu32 
            Caption         =   "員工貸款償還明細"
            Index           =   7
         End
         Begin VB.Menu mnu32 
            Caption         =   "員工差旅房租技術/證照津貼明細表"
            Index           =   8
         End
         Begin VB.Menu mnu32 
            Caption         =   "個人加班紀錄列印"
            Index           =   9
         End
         Begin VB.Menu mnu32 
            Caption         =   "個人出缺勤明細表"
            Index           =   10
         End
         Begin VB.Menu mnu32 
            Caption         =   "員工補充健保費查詢及列印"
            Index           =   11
         End
         Begin VB.Menu mnu32 
            Caption         =   "其他所得人補充保費查詢及列印"
            Index           =   12
         End
         Begin VB.Menu mnu32 
            Caption         =   "每月薪資異動金額明細"
            Index           =   13
         End
      End
      Begin VB.Menu mnu3 
         Caption         =   "不定期"
         Index           =   3
         Begin VB.Menu mnu33 
            Caption         =   "同仁婚喪互助明細表"
            Index           =   1
         End
         Begin VB.Menu mnu33 
            Caption         =   "薪資調薪表"
            Index           =   2
         End
         Begin VB.Menu mnu33 
            Caption         =   "敘薪/換敘通知單"
            Index           =   3
         End
         Begin VB.Menu mnu33 
            Caption         =   "薪資扣繳表"
            Index           =   4
         End
         Begin VB.Menu mnu33 
            Caption         =   "薪資扣繳表－人事用"
            Index           =   5
         End
         Begin VB.Menu mnu33 
            Caption         =   "地址條列印(寄扣單已不用)"
            Index           =   7
         End
         Begin VB.Menu mnu33 
            Caption         =   "端午,中秋代金入帳明細"
            Index           =   8
         End
         Begin VB.Menu mnu33 
            Caption         =   "例外扣繳項目員工名單"
            Index           =   9
         End
         Begin VB.Menu mnu33 
            Caption         =   "加班費明細表"
            Index           =   10
         End
         Begin VB.Menu mnu33 
            Caption         =   "福利金查詢及列印"
            Index           =   11
         End
         Begin VB.Menu mnu33 
            Caption         =   "智權人員薪點表"
            Index           =   12
         End
      End
      Begin VB.Menu mnu3 
         Caption         =   "互助會"
         Index           =   4
         Begin VB.Menu mnu34 
            Caption         =   "互助會名單"
            Index           =   1
         End
         Begin VB.Menu mnu34 
            Caption         =   "互助會得標金額明細"
            Index           =   2
         End
      End
      Begin VB.Menu mnu3 
         Caption         =   "年終獎金"
         Index           =   5
         Begin VB.Menu mnu35 
            Caption         =   "員工年終獎金明細"
            Index           =   1
         End
         Begin VB.Menu mnu35 
            Caption         =   "年終獎金發放明細"
            Index           =   2
         End
         Begin VB.Menu mnu35 
            Caption         =   "年終獎金入帳明細"
            Index           =   3
         End
         Begin VB.Menu mnu35 
            Caption         =   "未休假代金明細"
            Index           =   4
         End
         Begin VB.Menu mnu35 
            Caption         =   "年度特殊功績獎金"
            Index           =   5
         End
      End
      Begin VB.Menu mnu3 
         Caption         =   "扣繳憑單"
         Index           =   6
         Begin VB.Menu mnu36 
            Caption         =   "各類所得申報明細"
            Index           =   1
         End
         Begin VB.Menu mnu36 
            Caption         =   "扣繳憑單套印"
            Index           =   2
         End
         Begin VB.Menu mnu36 
            Caption         =   "勞、健、補充保費扣繳證明書"
            Index           =   3
         End
         Begin VB.Menu mnu36 
            Caption         =   "員工年度所得統計 (非申報數)"
            Index           =   4
         End
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "年終獎金作業"
      Index           =   4
      Begin VB.Menu mnu4 
         Caption         =   "基準月數資料"
         Index           =   1
      End
      Begin VB.Menu mnu4 
         Caption         =   "年終考績資料"
         Index           =   2
      End
      Begin VB.Menu mnu4 
         Caption         =   "特殊功績獎金輸入"
         Index           =   3
      End
      Begin VB.Menu mnu4 
         Caption         =   "計算年終獎金(試算)"
         Index           =   4
      End
      Begin VB.Menu mnu4 
         Caption         =   "特殊功績獎金清除"
         Index           =   5
      End
      Begin VB.Menu mnu4 
         Caption         =   "計算年終獎金"
         Index           =   6
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "扣繳憑單作業"
      Index           =   5
      Begin VB.Menu mnu5 
         Caption         =   "其他所得人基本資料"
         Index           =   1
      End
      Begin VB.Menu mnu5 
         Caption         =   "其他各類所得資料"
         Index           =   2
      End
      Begin VB.Menu mnu5 
         Caption         =   "其他各類所得資料(平日)"
         Index           =   3
      End
      Begin VB.Menu mnu5 
         Caption         =   "股利及退職所得資料"
         Index           =   4
      End
      Begin VB.Menu mnu5 
         Caption         =   "各類所得轉入扣繳憑單－所得資料"
         Index           =   5
      End
      Begin VB.Menu mnu5 
         Caption         =   "各類所得轉入媒體申報套裝軟體"
         Index           =   6
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "互助會作業"
      Index           =   6
      Begin VB.Menu mnu6 
         Caption         =   "互助會基本資料"
         Index           =   1
      End
      Begin VB.Menu mnu6 
         Caption         =   "互助會得標名單"
         Index           =   2
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "檔案維護"
      Index           =   7
      Begin VB.Menu mnu7 
         Caption         =   "基本資料"
         Index           =   1
         Begin VB.Menu mnu71 
            Caption         =   "薪資基本資料"
            Index           =   1
         End
         Begin VB.Menu mnu71 
            Caption         =   "公司基本資料"
            Index           =   2
         End
         Begin VB.Menu mnu71 
            Caption         =   "所得稅率表"
            Index           =   3
         End
         Begin VB.Menu mnu71 
            Caption         =   "勞保投保等級維護"
            Index           =   4
         End
         Begin VB.Menu mnu71 
            Caption         =   "健保投保等級維護"
            Index           =   5
         End
         Begin VB.Menu mnu71 
            Caption         =   "勞退投保等級維護"
            Index           =   6
         End
         Begin VB.Menu mnu71 
            Caption         =   "勞保勞退健保費率資料"
            Index           =   7
         End
         Begin VB.Menu mnu71 
            Caption         =   "其他所得/扣款代號"
            Index           =   8
         End
         Begin VB.Menu mnu71 
            Caption         =   "執行業務業別代號"
            Index           =   9
         End
         Begin VB.Menu mnu71 
            Caption         =   "勞保補助類別"
            Index           =   10
         End
         Begin VB.Menu mnu71 
            Caption         =   "健保補助類別"
            Index           =   11
         End
      End
      Begin VB.Menu mnu7 
         Caption         =   "其他資料"
         Index           =   2
         Begin VB.Menu mnu72 
            Caption         =   "每月薪資資料"
            Index           =   1
         End
         Begin VB.Menu mnu72 
            Caption         =   "年終獎金維護"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnuTitle 
      Caption         =   "系統"
      Index           =   15
      Begin VB.Menu mnu00 
         Caption         =   "切換連線"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnu00 
         Caption         =   "說明"
         Index           =   1
         Begin VB.Menu mnu151 
            Caption         =   "關於"
            Index           =   0
         End
      End
      Begin VB.Menu mnu00 
         Caption         =   "結束"
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
   Begin VB.Menu mnuDML 
      Caption         =   "查維護紀錄"
      Index           =   0
      Visible         =   0   'False
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
Option Explicit

Dim WithEvents eventConn As ADODB.Connection
Attribute eventConn.VB_VarHelpID = -1
Public bolReOpen As Boolean
'intPCaseKind分案之系統分類，intPWhere 0國內  1國外CF  2國外FC
Public intPCaseKind As Integer, intPWhere As Integer
Public m_wasMaximized As Boolean 'Added by Morgan 畫面最小化後判斷原來是否為最大化用


'控制連線閒置超過30分鐘自動關閉程式
Private Sub eventConn_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
   tmrConnect.Tag = 0
End Sub

Private Sub SwitchMenu(Optional bolEnable As Boolean = True)
Dim mnuTmp As Menu
   
   For Each mnuTmp In mnuTitle
      If mnuTmp.Index <> 0 Then mnuTmp.Enabled = bolEnable
   Next
   If bolEnable = False Then Toolbar1.Visible = False
End Sub

Private Sub CloseAllChild()
Dim frmTemp As Form
   
   For Each frmTemp In Forms
      If frmTemp.Name <> "mdiMain" Then Unload frmTemp
   Next
End Sub

Private Sub ReConnect()

   Timer1.Enabled = True
   Timer1.Interval = 100
   tmrConnect.Tag = 0
   
End Sub

Private Sub MDIForm_Activate()

'Added by Morgan 2012/1/5 視窗原為最大化,改指定大小及位置以方便其他應用程式操作
Static bolFormIsSet As Boolean
If bolFormIsSet = False Then
   PUB_InitFormPos Me
   bolFormIsSet = True
End If
'end 2012/1/5

End Sub

Private Sub MDIForm_Resize()
   'Added by Morgan 2011/12/14 紀錄是否為最大化狀態
   If Me.WindowState = 2 Then
      m_wasMaximized = True
   ElseIf Me.WindowState = 0 Then
      m_wasMaximized = False
   End If
End Sub

Private Sub mnu1_Click(Index As Integer)
    ToolHide
    Select Case Index
    Case 1   '其他所得/扣款資料
        If CheckUse("frm170002", strExec) = False Then
            Exit Sub
        End If
        frm170002.Show
    Case 2   '同仁其他給付資料
        If CheckUse("frm17003", strExec) = False Then
            Exit Sub
        End If
        frm170003.Show
    Case 3   '每月獎金資料
        If CheckUse("frm170004", strExec) = False Then
            Exit Sub
        End If
        frm170004.Show
    Case 4   '薪資異動資料
        If CheckUse("frm170005", strExec) = False Then
            Exit Sub
        End If
        frm170005.Show
    Case 5   '員工借支資料
        If CheckUse("frm170006", strExec) = False Then
            Exit Sub
        End If
        frm170006.Show
    Case 6   '員工貸款資料
        If CheckUse("frm170007", strExec) = False Then
            Exit Sub
        End If
        frm170007.Show
    Case 7   '兼職人員每月工作時數輸入
        If CheckUse("frm170008", strExec) = False Then
            Exit Sub
        End If
        frm170008.Show
    'Added by Morgan 2023/11/1
    Case 8   '尾牙摸彩、年資、全勤獎金維護
        If CheckUse("frm170032", strExec) = False Then
            Exit Sub
        End If
        frm170032.Show
    'end 2023/11/1
    Case Else
    End Select
End Sub

Private Sub mnu2_Click(Index As Integer)
    ToolHide
    Select Case Index
    Case 1   '每月薪資計算
        If CheckUse("frm170101", strExec) = False Then
            Exit Sub
        End If
        frm170101.Show
        
    'Removed by Morgan 2025/7/29 114/7/28起廢止婚喪互助辦法
    'Case 2   '婚喪互助扣款計算
    '    If CheckUse("frm170102", strExec) = False Then
    '        Exit Sub
    '    End If
    '    frm170102.Show
    
    Case 3   '勞健保保費重新計算
        If CheckUse("frm170103", strExec) = False Then
            Exit Sub
        End If
        frm170103.Show
    Case 4   '銀行入帳媒體作業
        If CheckUse("frm170104", strExec) = False Then
            Exit Sub
        End If
        frm170104.Show
    Case 5   '補充保費繳納作業
        If CheckUse("frm170105", strExec) = False Then
            Exit Sub
        End If
        frm170105.Show
    Case 6   '補充保費明細申報檔
        If CheckUse("frm170106", strExec) = False Then
            Exit Sub
        End If
        frm170106.Show
    'Added by Morgan 2023/11/30
    Case 7 '福利金轉檔
        If CheckUse("frm170109", strExec) = False Then
            Exit Sub
        End If
        frm170109.Show
    Case Else
    End Select
End Sub

Private Sub mnu31_Click(Index As Integer)
    ToolHide
    Select Case Index
    Case 1   '每月薪資資料查詢
        If CheckUse("frm170304", strExec) = False Then
            Exit Sub
        End If
        frm170304.Show
    Case 2   '個人勞退自提明細表   2009/7/6 add by sonia
        If CheckUse("frm170232", strExec) = False Then
            Exit Sub
        End If
        frm170232.Show
    Case 3   '薪資基本資料檢核表
        If CheckUse("frm170211", strExec) = False Then
            Exit Sub
        End If
        frm170211.Show
    Case 4   '人事薪資異動檢核表
        If CheckUse("frm170231", strExec) = False Then
            Exit Sub
        End If
        frm170231.Show
    Case 5   '員工應稅薪資檢核表
        If CheckUse("frm170219", strExec) = False Then
            Exit Sub
        End If
        frm170219.Show
    Case 6   '其他各類所得資料檢核表
        If CheckUse("frm170228", strExec) = False Then
            Exit Sub
        End If
        frm170228.Show
    Case 7   '股利及退職所得資料檢核表
        If CheckUse("frm170229", strExec) = False Then
            Exit Sub
        End If
        frm170229.Show
    Case Else
    End Select
End Sub

Private Sub mnu32_Click(Index As Integer)
    ToolHide
    Select Case Index
    Case 1   '員工薪資明細表
        If CheckUse("frm170207", strExec) = False Then
            Exit Sub
        End If
        frm170207.Show
    Case 2   '員工薪資單
        If CheckUse("frm170208", strExec) = False Then
            Exit Sub
        End If
        frm170208.Show
    Case 3   '薪資所得稅明細表
        If CheckUse("frm170220", strExec) = False Then
            Exit Sub
        End If
        frm170220.Show
    Case 4   '薪資媒體入帳明細
        If CheckUse("frm170209", strExec) = False Then
            Exit Sub
        End If
        frm170209.Show
    Case 5   '薪資媒體轉帳遞送單及員工清冊
        If CheckUse("frm170230", strExec) = False Then
            Exit Sub
        End If
        frm170230.Show
    Case 6   '員工勞健保保費及勞退自提明細表
        If CheckUse("frm170210", strExec) = False Then
            Exit Sub
        End If
        frm170210.Show
    Case 7   '員工貸款償還明細
        If CheckUse("frm170212", strExec) = False Then
            Exit Sub
        End If
        frm170212.Show
    Case 8   '員工差旅房租技術/證照津貼明細表
        If CheckUse("frm170215", strExec) = False Then
            Exit Sub
        End If
        frm170215.Show
    Case 9   '個人加班紀錄列印
        If CheckUse("frm160107", strExec) = False Then
            Exit Sub
        End If
        frm160107.Show
    Case 10  '個人出缺勤明細表
        If CheckUse("frm160113", strExec) = False Then
            Exit Sub
        End If
        frm160113.Show
    'Added by Morgan 2013/2/27
    Case 11 '員工補充健保費查詢及列印
        If CheckUse("frm170234", strExec) = False Then
            Exit Sub
        End If
        frm170234.Show
    'Added by Morgan 2013/3/1
    Case 12 '其他所得人補充保費查詢及列印
        If CheckUse("frm170235", strExec) = False Then
            Exit Sub
        End If
        frm170235.Show
        
    'Added by Morgan 2023/6/15
    Case 13 '每月薪資異動金額明細
        If CheckUse("frm170242", strExec) = False Then
            Exit Sub
        End If
        frm170242.Show
        
    Case Else
    End Select
End Sub

Private Sub mnu33_Click(Index As Integer)
    ToolHide
    Select Case Index
    Case 1   '同仁婚喪互助明細表
        If CheckUse("frm170206", strExec) = False Then
            Exit Sub
        End If
        frm170206.Show
    Case 2   '薪資調薪表
        If CheckUse("frm170202", strExec) = False Then
            Exit Sub
        End If
        frm170202.Show
    Case 3   '敘薪/換敘通知單
        If CheckUse("frm170201", strExec) = False Then
            Exit Sub
        End If
        frm170201.Show
    Case 4   '薪資扣繳表
        If CheckUse("frm170203", strExec) = False Then
            Exit Sub
        End If
        frm170203.Show
    Case 5   '薪資扣繳表－人事用
        If CheckUse("frm170226", strExec) = False Then
            Exit Sub
        End If
        frm170226.Show
'    Case 6   '員工歷年薪資明細表  2011/1/18發現還沒寫,辜說不要寫了
'        If CheckUse("frm170204", strExec) = False Then
'            Exit Sub
'        End If
'        frm170204.Show
    Case 7   '地紙條列印
        If CheckUse("frm170205", strExec) = False Then
            Exit Sub
        End If
        frm170205.Show
    Case 8   '端午,中秋代金入帳明細
        If CheckUse("frm170225", strExec) = False Then
            Exit Sub
        End If
        frm170225.Show
    Case 9   '例外扣繳項目員工名單
        If CheckUse("frm170224", strExec) = False Then
            Exit Sub
        End If
        frm170224.Show
    Case 10   '加班費明細表 'Added by Morgan 2012/6/19
        If CheckUse("frm170233", strExec) = False Then
            Exit Sub
        End If
        frm170233.Show
    'Added by Morgan 2023/11/23
    Case 11   '福利金查詢及列印
        If CheckUse("frm170243", strExec) = False Then
            Exit Sub
        End If
        frm170243.Show
    'end 2023/11/23
    'Added by Sindy 2024/3/19
    Case 12   '智權人員薪點表
        If CheckUse("frm170244", strExec) = False Then
            Exit Sub
        End If
        frm170244.Show
    'end 2024/3/19
    Case Else
    End Select
End Sub

Private Sub mnu34_Click(Index As Integer)
    ToolHide
    Select Case Index
    Case 1   '互助會名單
        If CheckUse("frm170214", strExec) = False Then
            Exit Sub
        End If
        frm170214.Show
    Case 2   '互助會得標金額明細
        If CheckUse("frm170213", strExec) = False Then
            Exit Sub
        End If
        frm170213.Show
    Case Else
    End Select
End Sub

Private Sub mnu35_Click(Index As Integer)
    ToolHide
    Select Case Index
    Case 1   '員工年終獎金明細表
        If CheckUse("frm170216", strExec) = False Then
            Exit Sub
        End If
        frm170216.Show
    Case 2   '年終獎金發放明細
        If CheckUse("frm170217", strExec) = False Then
            Exit Sub
        End If
        frm170217.Show
    Case 3   '年終獎金入帳明細
        If CheckUse("frm170218", strExec) = False Then
            Exit Sub
        End If
        frm170218.Show
    Case 4   '未休假代金明細
        If CheckUse("frm170227", strExec) = False Then
            Exit Sub
        End If
        frm170227.Show
    'Add By Sindy 2020/11/16
    Case 5   '年度特殊功績獎金
        If CheckUse("frm170240", strExec) = False Then
            Exit Sub
        End If
        frm170240.Show
    Case Else
    End Select
End Sub

Private Sub mnu36_Click(Index As Integer)
    ToolHide
    Select Case Index
    Case 1   '各類所得申報明細
        If CheckUse("frm170223", strExec) = False Then
            Exit Sub
        End If
        frm170223.Show
    Case 2   '扣繳憑單套印
        If CheckUse("frm170221", strExec) = False Then
            Exit Sub
        End If
        frm170221.Show
    Case 3   '勞、健、補充保費扣繳證明書
        If CheckUse("frm170222", strExec) = False Then
            Exit Sub
        End If
        frm170222.Show
    'Add By Sindy 2021/2/2
    Case 4   '員工年度所得統計 (非申報數)
        If CheckUse("frm170241", strExec) = False Then
            Exit Sub
        End If
        frm170241.Show
    Case Else
    End Select
End Sub

Private Sub mnu4_Click(Index As Integer)
    ToolHide
    Select Case Index
    Case 1   '基準月數輸入
        If CheckUse("frm170024", strExec) = False Then
            Exit Sub
        End If
        frm170024.Show
    Case 2   '年終考績輸入
        If CheckUse("frm170020", strExec) = False Then
            Exit Sub
        End If
        frm170020.Show
    Case 3   '特殊功績獎金輸入
        If CheckUse("frm170025", strExec) = False Then
            Exit Sub
        End If
        frm170025.Show
    
    'Added by Morgan 2023/12/11
    Case 4   '計算年終獎金(試算)
        If CheckUse("frm170302", strExec) = False Then
            Exit Sub
        End If
        If PUB_CheckFormExist("frm170302") = False Then
            frm170302.m_bolIsTrial = True
            frm170302.Show
        End If
        
    Case 5  '特殊功績獎金清除
        If CheckUse("frm170026", strExec) = False Then
            Exit Sub
        End If
        frm170026.Show
        
    Case 6   '計算年終獎金
        If CheckUse("frm170302", strExec) = False Then
            Exit Sub
        End If
        If PUB_CheckFormExist("frm170302") = False Then 'Added by Morgan 2023/12/11
            frm170302.Show
        End If
    'end 2023/12/11
    Case Else
    End Select
End Sub

Private Sub mnu5_Click(Index As Integer)
    ToolHide
    Select Case Index
    Case 1   '其他所得人基本資料
        If CheckUse("frm170009", strExec) = False Then
            Exit Sub
        End If
        frm170009.Show
    Case 2   '其他各類所得資料
        If CheckUse("frm170027", strExec) = False Then
            Exit Sub
        End If
        frm170027.Show
        
    Case 3   '其他各類所得資料(平日)
        If CheckUse("frm170031", strExec) = False Then
            Exit Sub
        End If
        frm170031.Show
        
    Case 4   '股利及退職所得資料
        If CheckUse("frm170028", strExec) = False Then
            Exit Sub
        End If
        frm170028.Show
    Case 5   '各類所得轉入扣繳憑單－所得資料
        If CheckUse("frm170305", strExec) = False Then
            Exit Sub
        End If
        frm170305.Show
    Case 6   '各類所得轉入媒體申報套裝軟體
        If CheckUse("frm170306", strExec) = False Then
            Exit Sub
        End If
        frm170306.Show
    Case Else
    End Select
End Sub

Private Sub mnu6_Click(Index As Integer)
    ToolHide
    Select Case Index
    Case 1   '互助會基本資料
        If CheckUse("frm170011", strExec) = False Then
            Exit Sub
        End If
        frm170011.Show
    'Modified by Morgan 2022/3/3 改標題--婧瑄
    Case 2   '互助會得標名單
        If CheckUse("frm170012", strExec) = False Then
            Exit Sub
        End If
        frm170012.Show
    Case Else
    End Select
End Sub

Private Sub mnu71_Click(Index As Integer)
    ToolHide
    Select Case Index
    Case 1   '薪資基本資料
        If CheckUse("frm170001", strExec) = False Then
            Exit Sub
        End If
        frm170001.Show
    Case 2   '公司基本資料
        If CheckUse("frm170014", strExec) = False Then
            Exit Sub
        End If
        frm170014.Show
    Case 3   '所得稅率表
        If CheckUse("frm170015", strExec) = False Then
            Exit Sub
        End If
        frm170015.Show
    Case 4   '勞保投保等級維護
        If CheckUse("frm170016", strExec) = False Then
            Exit Sub
        End If
        ProSysState = "L"
        frm170016.Show
    Case 5   '健保投保等級維護
        If CheckUse("frm170016H", strExec) = False Then
            Exit Sub
        End If
        ProSysState = "H"
        frm170016.Show
    Case 6   '勞退投保等級維護
        If CheckUse("frm170016R", strExec) = False Then
            Exit Sub
        End If
        ProSysState = "R"
        frm170016.Show
    Case 7   '勞保勞退健保費率資料
        If CheckUse("frm170018", strExec) = False Then
            Exit Sub
        End If
        frm170018.Show
    Case 8   '其他所得/扣款代號
        If CheckUse("frm170019", strExec) = False Then
            Exit Sub
        End If
        frm170019.Show
    Case 9   '執行業務業別代號
        If CheckUse("frm160011G", strExec) = False Then
            Exit Sub
        End If
        ProSysState = "G"
        frm160011.Show
        
    'Add by Morgan 2009/6/24
    Case 10   '勞保補助類別
        If CheckUse("frm170029", strExec) = False Then
            Exit Sub
        End If
        frm170029.Show
        
    Case 11   '健保補助類別
        If CheckUse("frm170030", strExec) = False Then
            Exit Sub
        End If
        frm170030.Show
    'end 2009/6/24
    
    Case Else
    End Select
End Sub

Private Sub mnu72_Click(Index As Integer)
    ToolHide
    Select Case Index
    Case 1   '每月薪資資料
        If CheckUse("frm170021", strExec) = False Then
            Exit Sub
        End If
        frm170021.Show
    Case 2  '年終獎金維護
        If CheckUse("frm170022", strExec) = False Then
            Exit Sub
        End If
        frm170022.Show
    Case Else
    End Select
End Sub

Private Sub mnuDML_Click(Index As Integer)
    frmDML.Show   '基本資料維護紀錄
End Sub

'控制不可拷貝畫面
Private Sub Timer3_Timer()
   'Added by Morgan 2024/8/8 定時執行一次語法以確保跨網段連線時網路不會被切斷
   Static dtNow As Date
      
On Error Resume Next '若有其他軟體也在使用剪貼簿時會發生521(無法開啟剪貼簿)的錯誤(Ex.Word開啟剪貼簿並擷取畫面)
   
   If Now > dtNow Then
      dtNow = DateAdd("n", cntAutoQueryInterval, Now)
      ClsLawReadRstMsg 1, "select * from dual"
   End If
   'end 2024/8/8
   
   '電腦中心的不管
   If Pub_StrUserSt03 = "M51" Or Pub_Can_Copy_Pic = True Then Exit Sub
   '圖檔才清
   If Clipboard.GetFormat(1) = False And Clipboard.GetFormat(2) = True And Clipboard.GetFormat(3) = False Then
       Clipboard.Clear
   End If
End Sub
'2009/2/3 CANCEL BY SONIA
''控制連線閒置超過10分鐘自動離線
'Private Sub tmrConnect_Timer()
'   tmrConnect.Tag = tmrConnect.Tag + 1
'   If tmrConnect.Tag = 10 Then
'      Timer1.Enabled = False
'      bolReOpen = False
'      frmReopen.Show vbModal, Me
'      If bolReOpen = True Then
'         Call ReConnect
'      Else
'         Call mnu15_Click(2)
'      End If
'   End If
'End Sub
'2009/2/3 END

Private Sub MDIForm_Load()
'Add by Morgan 2003/12/23
'控制連線閒置超過30分鐘自動關閉程式
If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") = 0 Then
   Set eventConn = cnnConnection
   tmrConnect.Interval = 60000
End If
       DisableControl Me
Dim strSysKind As String
Dim lngValue, lngBufferSize As Long, intCounter As Integer
Dim strUserId As String * 10, strLocalId As String

    '若登入成功
    If pub_str_LoginSucceeded = "1" Then
        Me.Timer1.Interval = 100
       strSysKind = GetSystemKindByNick
       '可以查詢維護紀錄
       If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Or Pub_StrUserSt03 = "M51" Then
            mnuDML(0).Visible = True
       Else
            mnuDML(0).Visible = False
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

Private Sub MDIForm_Unload(Cancel As Integer)
Dim frm As Form
   '關閉尚未關閉的子視窗
   For Each frm In Forms
       If frm.Name <> mdiMain.Name Then
           Unload frm
       End If
   Next
   
   PUB_AddAuditLog AL_登出 'Added by Morgan 2025/7/31
   
   Set mdiMain = Nothing
End Sub
'切換連線選擇
Private Sub mnu00_Click(Index As Integer)
   Select Case Index
      Case 0
         If PUB_Connect2DB(True) = False Then
            Unload Me
         End If
      Case 2
         Unload Me
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

Private Sub Timer1_Timer()
Dim frm As Form
   '控制"視窗"Menu
   MenuForFormControl
   StatusBar1.Panels.Item(4).Text = time
   
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
