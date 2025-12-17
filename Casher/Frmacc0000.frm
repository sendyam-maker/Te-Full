VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm Frmacc0000 
   BackColor       =   &H8000000C&
   Caption         =   "分所財務"
   ClientHeight    =   4800
   ClientLeft      =   5180
   ClientTop       =   3460
   ClientWidth     =   7840
   Icon            =   "Frmacc0000.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  '最大化
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '對齊表單下方
      Height          =   384
      Left            =   0
      TabIndex        =   1
      Top             =   4416
      Width           =   7836
      _ExtentX        =   13829
      _ExtentY        =   670
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '對齊表單上方
      Height          =   560
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7840
      _ExtentX        =   13829
      _ExtentY        =   988
      ButtonWidth     =   494
      ButtonHeight    =   882
      Appearance      =   1
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3630
      Top             =   2130
      _ExtentX        =   988
      _ExtentY        =   988
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Timer tmrSalary 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1395
      Top             =   3000
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   900
      Top             =   915
   End
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   480
   End
   Begin VB.Timer Timer2 
      Left            =   480
      Top             =   900
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
   Begin VB.Menu Main1 
      Caption         =   "收款作業(&I)"
      Begin VB.Menu Main1_1 
         Caption         =   "收款輸入作業"
         Visible         =   0   'False
      End
      Begin VB.Menu Main1_2 
         Caption         =   "簽收作業"
      End
      Begin VB.Menu Main1_3 
         Caption         =   "客戶電匯資料維護及查詢"
      End
      Begin VB.Menu Main1_4 
         Caption         =   "客戶財務EMail資料維護"
      End
      Begin VB.Menu Main1_5 
         Caption         =   "智權人員繳款確認"
      End
   End
   Begin VB.Menu Main6 
      Caption         =   "分所收據"
      Begin VB.Menu Main6_1 
         Caption         =   "收據抬頭基本資料維護"
      End
      Begin VB.Menu Main6_4 
         Caption         =   "收據抬頭修改"
      End
      Begin VB.Menu Main6_2 
         Caption         =   "國內應收待處理作業"
      End
      Begin VB.Menu Main6_3 
         Caption         =   "收據列印"
         Begin VB.Menu Main6_3_1 
            Caption         =   "收據列印"
            Index           =   1
         End
         Begin VB.Menu Main6_3_1 
            Caption         =   "補開收據列印"
            Index           =   2
         End
         Begin VB.Menu Main6_3_1 
            Caption         =   "請款單及發票列印"
            Index           =   3
         End
         Begin VB.Menu Main6_3_1 
            Caption         =   "補開請款單及發票列印"
            Index           =   4
         End
      End
   End
   Begin VB.Menu Main2 
      Caption         =   "帳目查詢(&E)"
      Begin VB.Menu Main2_1 
         Caption         =   "分所收款資料查詢"
         Visible         =   0   'False
      End
      Begin VB.Menu Main2_14 
         Caption         =   "分所收款資料查詢-智權人員繳款"
      End
      Begin VB.Menu Main2_2 
         Caption         =   "收據資料查詢"
      End
      Begin VB.Menu Main2_12 
         Caption         =   "手開收據資料查詢"
      End
      Begin VB.Menu Main2_3 
         Caption         =   "客戶帳款查詢"
      End
      Begin VB.Menu Main2_4 
         Caption         =   "智權人員帳款查詢"
      End
      Begin VB.Menu Main2_5 
         Caption         =   "本所案號帳目查詢"
      End
      Begin VB.Menu Main2_11 
         Caption         =   "收款單號查詢"
      End
      Begin VB.Menu Main2_6 
         Caption         =   "智權人員點數查詢"
      End
      Begin VB.Menu Main2_7 
         Caption         =   "扣繳憑單查詢及列印"
      End
      Begin VB.Menu Main2_16 
         Caption         =   "每月提醒代填繳款書客戶明細"
      End
      Begin VB.Menu Main2_15 
         Caption         =   "會計師資料／客戶E-Mail資料查詢"
      End
      Begin VB.Menu Main2_8 
         Caption         =   "業績點數查詢"
      End
      Begin VB.Menu Main2_9 
         Caption         =   "暫收款查詢"
      End
      Begin VB.Menu Main2_10 
         Caption         =   "簽收資料查詢"
      End
      Begin VB.Menu Main2_13 
         Caption         =   "客戶應收帳款查詢"
      End
      Begin VB.Menu Main2_17 
         Caption         =   "科目分類帳查詢"
      End
      Begin VB.Menu Main2_18 
         Caption         =   "未繳款資料查詢與銀存核對"
      End
   End
   Begin VB.Menu Main3 
      Caption         =   "報表列印(&P)"
      Begin VB.Menu Main3_1 
         Caption         =   "分所每日收款明細表"
         Visible         =   0   'False
      End
      Begin VB.Menu Main3_8 
         Caption         =   "分所每日收款明細表-智權人員繳款"
      End
      Begin VB.Menu Main3_2 
         Caption         =   "分所智權人員收款明細表"
         Visible         =   0   'False
      End
      Begin VB.Menu Main3_9 
         Caption         =   "分所智權人員收款明細表-智權人員繳款"
      End
      Begin VB.Menu Main3_3 
         Caption         =   "客戶對帳單"
         Visible         =   0   'False
      End
      Begin VB.Menu Main3_4 
         Caption         =   "客戶帳款明細表"
         Visible         =   0   'False
      End
      Begin VB.Menu Main3_5 
         Caption         =   "智權人員帳款明細表"
      End
      Begin VB.Menu Main3_6 
         Caption         =   "銷帳退費明細表"
      End
      Begin VB.Menu Main3_7 
         Caption         =   "暫收款明細表"
      End
   End
   Begin VB.Menu Main4 
      Caption         =   "收據作業"
      Begin VB.Menu Main4_1 
         Caption         =   "基本資料"
         Begin VB.Menu Main4_1_1 
            Caption         =   "收據開立作業"
            Index           =   1
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "收據開立作業-批次"
            Index           =   2
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "收據抬頭修改"
            Index           =   3
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "收據/請款單作廢作業"
            Index           =   4
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "收據抬頭基本資料維護"
            Index           =   5
         End
         Begin VB.Menu Main4_1_1 
            Caption         =   "國內應收待處理作業"
            Index           =   6
         End
      End
      Begin VB.Menu Main4_2 
         Caption         =   "查詢作業"
         Begin VB.Menu Main4_2_1 
            Caption         =   "收文與收據資料檢核查詢"
            Index           =   1
         End
         Begin VB.Menu Main4_2_1 
            Caption         =   "本所案號帳目查詢"
            Index           =   2
         End
      End
      Begin VB.Menu Main_4_3 
         Caption         =   "報表列印"
         Begin VB.Menu Main4_3_1 
            Caption         =   "收據列印"
            Index           =   1
         End
         Begin VB.Menu Main4_3_1 
            Caption         =   "補開收據列印"
            Index           =   2
         End
         Begin VB.Menu Main4_3_1 
            Caption         =   "請款單及發票列印"
            Index           =   3
         End
         Begin VB.Menu Main4_3_1 
            Caption         =   "補開請款單及發票列印"
            Index           =   4
         End
      End
   End
   Begin VB.Menu Main5 
      Caption         =   "共同查詢(&Q)"
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
      Caption         =   "一般作業(&W)"
      Index           =   23
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
         Caption         =   "分所銀行入帳媒體作業"
         Index           =   3
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
         Caption         =   "圖書借閱資料查詢 "
         Index           =   7
      End
      Begin VB.Menu mnu23 
         Caption         =   "行事曆提醒通知"
         Index           =   8
      End
      Begin VB.Menu mnu23 
         Caption         =   "風險檢查對象資料維護"
         Index           =   9
      End
   End
   Begin VB.Menu mnuChUser 
      Caption         =   "更改使用者"
   End
   Begin VB.Menu Main7 
      Caption         =   "系統(&S)"
      Begin VB.Menu Main7_0 
         Caption         =   "切換連線"
         Visible         =   0   'False
      End
      Begin VB.Menu Main7_1 
         Caption         =   "結束"
      End
   End
End
Attribute VB_Name = "Frmacc0000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/28 日期欄已修改
Option Explicit

Private Declare Function getUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Dim intCounter As Integer
Public m_ChkIsOpenFrm180203 As Boolean 'Add By Sindy 2013/7/8
Public str中所收據人員 As String 'Added by Lydia 2020/03/26


Private Sub Main1_1_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc7100", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc7100.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main1_2_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc41e0", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc41e0.Show
   Me.MousePointer = vbDefault
End Sub

'Add By Sindy 2012/9/17
Private Sub Main1_3_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc11n1", strExec) = False Then
      Exit Sub
   End If
   'tool1_enabled
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc11n1.Show
   Me.MousePointer = vbDefault
End Sub

'add by sonia 2014/11/14 客戶/代理人財務EMail資料維護(分所用改顯示客戶財務EMail資料維護)
Private Sub Main1_4_Click()
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

Private Sub Main1_5_Click()
   If CheckUse("Frmacc7150", strExec) = False Then
      Exit Sub
   End If
   Toolbar1.Visible = False
   StatusBar1.Visible = False
   MenuDisabled
   Me.MousePointer = vbHourglass
   Frmacc7150.Show
   Me.MousePointer = vbDefault
End Sub

'Added by Morgan 2025/5/26
'科目分類帳查詢
Private Sub Main2_17_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc4220", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc4220.Show
   Me.MousePointer = vbDefault
   
End Sub

'Added by Morgan 2025/6/17
'未繳款資料查詢與銀存核對
Private Sub Main2_18_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc42b0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc42b0.Show
   Me.MousePointer = vbDefault
   
End Sub

'Add By Sindy 2015/7/9 收據抬頭基本資料維護
Private Sub Main6_1_Click()
   If CheckUse("Frmacc11p0", strExec) = False Then
      Exit Sub
   End If
   tool1_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc11p0.ProState = "2"  'add by sonia 2025/5/22 權限: 1.全所 2.該所
   Frmacc11p0.Show
   Me.MousePointer = vbDefault
End Sub

'add by sonia 2023/5/12 國內應收待處理作業(從收據作業複製出來)
Private Sub Main6_2_Click()
   'add by sonia 2025/5/22 發現沒控制
   If CheckUse("Frmacc11r0", strExec) = False Then
      Exit Sub
   End If
   'end 2025/5/22
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc11r0.Show
   Me.MousePointer = vbDefault
End Sub
'end 2023/5/12

Private Sub Main1_Click()
    Toolbar1.Visible = True
    StatusBar1.Visible = True
End Sub

Private Sub Main2_1_Click() '分所收款資料查詢
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc7110", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc7110.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main2_14_Click() 'Add by Lydia 2014/10/3 分所收款資料查詢-智權人員繳款
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc7160", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc7160.Show
   Me.MousePointer = vbDefault
End Sub

'Add by Morgan 2005/4/11
Private Sub Main2_10_Click()
   Toolbar1.Visible = False
   StatusBar1.Visible = False
   MenuDisabled
   'Modified by Lydia 2017/01/26 是否需要輸入密碼
   'frm210106.Show
    If frm210106_1.setNextForm = "" Then
       frm210106.Show
    Else
       frm210106_1.setCaller frm210106
       frm210106_1.Show
    End If
End Sub
'Add by Morgan 2008/9/15
Private Sub Main2_12_Click()
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

'Add By Sindy 2013/1/9
'客戶應收帳款查詢
Private Sub Main2_13_Click()
   If CheckUse("frm210122", strExec) = False Then
      Exit Sub
   End If
   Toolbar1.Visible = False
   StatusBar1.Visible = False
   MenuDisabled
   frm210122.cmdEdit.Enabled = False
   frm210122.Show
End Sub

'Add By Sindy 2017/9/27 每月提醒代填繳款書客戶明細
Private Sub Main2_16_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc44w0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc44w0.Command1.Enabled = False
   Frmacc44w0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main2_2_Click()
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

Private Sub Main2_3_Click()
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

Private Sub Main2_4_Click()
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
Private Sub Main2_5_Click()
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

Private Sub Main2_11_Click()
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

Private Sub Main2_6_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc4250", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc4250.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main2_7_Click()
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
'Add by Morgan 2005/4/11
Private Sub Main2_8_Click()
   Toolbar1.Visible = False
   StatusBar1.Visible = False
   MenuDisabled
   frm210104.Show
End Sub
'Add by Morgan 2005/4/11
Private Sub Main2_9_Click()
   Toolbar1.Visible = False
   StatusBar1.Visible = False
   MenuDisabled
   frm210105.Show
End Sub

Private Sub Main2_Click()
    Toolbar1.Visible = True
    StatusBar1.Visible = True
End Sub

Private Sub Main3_1_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc7120", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc7120.Show
   Me.MousePointer = vbDefault
End Sub

'Add By Sindy 2021/5/21
Private Sub Main6_3_1_Click(Index As Integer)
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   
   Select Case Index
      Case 1 '收據列印
         'add by sonia 2025/5/22 發現沒控制
         If CheckUse("Frmacc1410", strExec) = False Then
            Exit Sub
         End If
         'end 2025/5/22
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1410.ProState = "2" '權限: 1.全所 2.該所 Add By Sindy 2021/5/21
         Frmacc1410.Show
         Me.MousePointer = vbDefault
      Case 2 '補開收據列印
         'add by sonia 2025/5/22 發現沒控制
         If CheckUse("Frmacc1420", strExec) = False Then
            Exit Sub
         End If
         'end 2025/5/22
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1420.ProState = "2" '權限: 1.全所 2.該所 Add By Sindy 2021/5/21
         Frmacc1420.Show
         'Add By Sindy 2021/7/29 出納系統從報表列印的補開收據列印進去時，
         '隱藏收據抬頭修改按鈕，但從收據作業->報表列印的補開收據列印進去時則要出現。
         Frmacc1420.Command2.Visible = False
         Me.MousePointer = vbDefault
      Case 3 '請款單及發票列印
         'add by sonia 2025/5/22 發現沒控制
         If CheckUse("Frmacc1610", strExec) = False Then
            Exit Sub
         End If
         'end 2025/5/22
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1610.ProState = "2" '權限: 1.全所 2.該所 Add By Sindy 2021/5/21
         Frmacc1610.Show
         Me.MousePointer = vbDefault
      Case 4 '補開請款單及發票列印
         'add by sonia 2025/5/22 發現沒控制
         If CheckUse("Frmacc1620", strExec) = False Then
            Exit Sub
         End If
         'end 2025/5/22
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1620.ProState = "2" '權限: 1.全所 2.該所 Add By Sindy 2021/5/21
         Frmacc1620.Show
         Me.MousePointer = vbDefault
   End Select
End Sub

Private Sub Main3_2_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc7130", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc7130.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_8_Click() 'Add by Lydia 2014/10/3 分所每日收款明細表-智權人員繳款
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc7170", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc7170.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_9_Click() 'Add by Lydia 2014/10/3 分所智權人員收款明細表-智權人員繳款(個人)
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc7180", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc7180.Show
   Me.MousePointer = vbDefault
End Sub

'Mark by Lydia 2024/11/01 欲收回分所報表列印程式
'Private Sub Main3_3_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc1440", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc1440.Show
'   Me.MousePointer = vbDefault
'End Sub

'Mark by Lydia 2024/11/01 欲收回分所報表列印程式
'Private Sub Main3_4_Click()
'   If strFormName <> MsgText(601) Then
'      Exit Sub
'   End If
'   If CheckUse("Frmacc1450", strExec) = False Then
'      Exit Sub
'   End If
'   tool3_enabled
'   Me.MousePointer = vbHourglass
'   MenuDisabled
'   Frmacc1450.Show
'   Me.MousePointer = vbDefault
'End Sub

'Mark by Lydia 2024/11/01 欲收回分所報表列印程式
'Modified by Lydia 2024/12/02 再開放智權人員請款明細表
Private Sub Main3_5_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc1460", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc1460.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_6_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc1490", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc1490.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_7_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   If CheckUse("Frmacc14a0", strExec) = False Then
      Exit Sub
   End If
   tool3_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc14a0.Show
   Me.MousePointer = vbDefault
End Sub

Private Sub Main3_Click()
    Toolbar1.Visible = True
    StatusBar1.Visible = True
End Sub

Private Sub Main5_1_Click(Index As Integer)
    Select Case Index
    Case 1
        If CheckUse("frm100101_1", strExec) = False Then
            Exit Sub
        End If
        frm100101_1.Show
    Case 2
        If CheckUse("frm100102_1", strExec) = False Then
            Exit Sub
        End If
        frm100102_1.Show
    Case 3
        If CheckUse("frm100120_1", strExec) = False Then
            Exit Sub
        End If
        frm100120_1.Show
    Case 4
        If CheckUse("frm100103_1", strExec) = False Then
            Exit Sub
        End If
        frm100103_1.Show
    Case 5
        If CheckUse("frm100104_1", strExec) = False Then
            Exit Sub
        End If
        frm100104_1.Show
    Case 6
        If CheckUse("frm100105_1", strExec) = False Then
            Exit Sub
        End If
        frm100105_1.Show
    Case 7
        If CheckUse("frm100106_1", strExec) = False Then
            Exit Sub
        End If
        frm100106_1.Show
    Case 8
        If CheckUse("frm100107_1", strExec) = False Then
            Exit Sub
        End If
        frm100107_1.Show
    Case 9
        If CheckUse("frm100108_1", strExec) = False Then
            Exit Sub
        End If
        frm100108_1.Show
    Case 10
        If CheckUse("frm100109_1", strExec) = False Then
            Exit Sub
        End If
        frm100109_1.Show
    Case 11
        If CheckUse("frm100110_1", strExec) = False Then
            Exit Sub
        End If
        frm100110_1.Show
    Case 12
        If CheckUse("frm100111_1", strExec) = False Then
            Exit Sub
        End If
        frm100111_1.Show
    Case 13
        If CheckUse("frm10011201_1", strExec) = False Then
            Exit Sub
        End If
        frm10011201_1.Show
    Case 14
        If CheckUse("frm10011202_1", strExec) = False Then
            Exit Sub
        End If
        frm10011202_1.Show
    Case 15
        If CheckUse("frm100113_1", strExec) = False Then
            Exit Sub
        End If
        frm100113_1.Show
    Case 16
        If CheckUse("frm100114_1", strExec) = False Then
            Exit Sub
        End If
        frm100114_1.Show
    Case 17
        If CheckUse("frm100115_1", strExec) = False Then
            Exit Sub
        End If
        frm100115_1.Show
    Case 18
        If CheckUse("frm100116_1", strExec) = False Then
            Exit Sub
        End If
        frm100116_1.Show
    Case 19
        If CheckUse("frm100117_1", strExec) = False Then
            Exit Sub
        End If
        frm100117_1.Show
    Case 20
        If CheckUse("frm100118_1", strExec) = False Then
            Exit Sub
        End If
        frm100118_1.Show
    Case 21
        If CheckUse("frm100119_1", strExec) = False Then
            Exit Sub
        End If
        frm100119_1.Show
    Case 22
        If CheckUse("frm100121_1", strExec) = False Then
            Exit Sub
        End If
        frm100121_1.Show
    '2005/8/2 CANCEL BY SONIA
    'Case 23 '業務收/發文量比較查詢
    '    If CheckUse("frm100122_1", strExec) = False Then
    '        Exit Sub
    '    End If
    '    frm100122_1.Show
      'add by nickc 2007/07/06
      Case 23 '客戶重新委任案件查詢列印
         If CheckUse("frm100126_1", strExec) Then
            frm100126_1.Show
         End If
      'Add By Sindy 2009/10/02
      Case 24 '優先權資料查詢
         If CheckUse("frm100127_1", strExec) Then
            frm100127_1.Show
         End If
      '2011/5/20 ADD BY SONIA
      Case 25  '庭期資料查詢
         'If CheckUse("frm072001", strExec) = True Then
            frm072001.Show
         'End If
    End Select
End Sub

Private Sub Main5_Click()
    Toolbar1.Visible = False
    StatusBar1.Visible = False
End Sub

'add by sonia 2023/5/26  收據抬頭修改 辜又說開放分所改該分所的資料
Private Sub Main6_4_Click()
   'add by sonia 2025/5/22 發現沒控制
   If CheckUse("Frmacc1140", strExec) = False Then
      Exit Sub
   End If
   'end 2025/5/22
   tool8_enabled
   Me.MousePointer = vbHourglass
   MenuDisabled
   Frmacc1140.ProState = "2" '權限: 1.全所 2.該所
   Frmacc1140.Show
   Me.MousePointer = vbDefault
End Sub
'end 2023/5/26

'Add by Morgan 2005/12/15 切換連線
Private Sub Main7_0_Click()
   If PUB_Connect2DB(True) = False Then
      Unload Me
   End If
End Sub

Private Sub Main7_1_Click()
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   'Modified by Morgan 2025/7/31
   'End
   Unload Me
   'end 2025/7/31
End Sub

Private Sub MDIForm_Activate()
   'Modify By Sindy 2025/11/3 改為共用函數
   Call MDIFormStarProc
End Sub

'Add By Sindy 2025/11/3
Public Sub SetTmpForm()
   Set Tmpfrm180201 = frm180201
   Set Tmpfrm180101 = frm180101
   Set Tmpfrm180203_1 = frm180203_1
   Set Tmpfrm160102 = frm160102
   Set Tmpfrm160018 = frm160018
   Set Tmpfrm010035_2 = frm010035_2
End Sub
'Add By Sindy 2011/10/7
Public Sub SysStartCallForm()
   '此函數在各系統一啟動時,因出缺勤待辦提示納入之故,共用會使用到,所以不可刪除
End Sub

'*************************************************
'  工具列按鈕及圖案設定
'
'*************************************************
Private Sub MDIForm_Load()
Dim lngValue, lngBufferSize As Long, intCount As Integer
Dim strUserId As String * 10, strLocalId As String
    
    If pub_str_LoginSucceeded = "1" Then
'       Show
'       Me.Enabled = False
'       Frmacc0002.Show
'       DoEvents
'       Unload Frmacc0002
'       Me.Enabled = True
       lngBufferSize = 10
       If strUserNum = "" Then
        '   lngValue = WNetGetUser(strLocalId, strUserId, lngBufferSize)
           lngValue = getUserName(strUserId, lngBufferSize)
           For intCount = 1 To 10
              If Asc(Mid(strUserId, intCount, 1)) = 0 Then
                 Exit For
              End If
              strUserNum = strUserNum & Mid(strUserId, intCount, 1)
              strUserNum = strUserNum
           Next intCount
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
        'Add By Cheng 2004/01/12
        '取得使用者所別
        pub_strUserOffice = PUB_GetST06(strUserNum)
        'End
        strSrvDate(1) = Format(ServerDate)
        strSrvDate(2) = Format(Val(strSrvDate(1)) - 19110000)
        'Add by Amy 2017/01/25
        If strSrvDate(1) >= 20170202 Then
          mnu23(7).Visible = True
        Else
          mnu23(7).Visible = False
        End If
        'Added by Lydia 2020/03/26 特別開放Casher的收據作業
        str中所收據人員 = Pub_GetSpecMan("中所收據人員")
        If (str中所收據人員 <> "" And InStr(str中所收據人員, strUserNum) > 0) Or PUB_GetST03(strUserNum) = "M51" Then
             Main4.Visible = True
        Else
             Main4.Visible = False
        End If
        'end 2020/03/26
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
    'add by nick 2004/08/20
    strUserDept = GetStaffDepartment(strUserNum)

'       'Added by Morgan 2016/1/22 薪資查詢測試
'       If strUserDept = "M51" Then
'         mnu23(6).Visible = True
'       Else
'         mnu23(6).Visible = False
'       End If
'       'end

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   PUB_AddAuditLog AL_登出 'Added by Morgan 2025/7/31
'   objOraDatabase.Close
'   objOraSession.Close
   adoTaie.Close
   Set Frmacc0000 = Nothing
End Sub

Private Sub mnu10_Click(Index As Integer)
'   ToolHide
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
'   ToolHide
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
'   ToolHide
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
'   ToolHide
   Select Case Index
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

'2013/3/6 add by sonia
Private Sub mnu23_Click(Index As Integer)
   Select Case Index
      'Modify By Sindy 2023/8/9 Mark,因南所入帳媒體也改在北所統一操作
'      Case 3 '分所銀行入帳媒體作業
'         If CheckUse("Frmacc7140", strExec) Then
'            Frmacc7140.Show
'         End If
      Case 7 '圖書借閱資料查詢 Add by Amy 2017/01/25
         frm010035.Show
         'Add by Amy 2017/02/03 判斷是否有圖書借閱記錄需簽核
        If GetLoanRecordApply = True Then
            frm010035.bolLoanRecordApply = True
            Call frm010035.cmdLoanRecord_Click
         End If
     'Added by Lydia 2020/01/15
     Case 8 '行事曆提醒通知
         frm060209.m_Role = "F41"
         frm060209.Show
      'Add by Amy 2024/02/01
      Case 9 '風險檢查對象資料維護
         frm12040163.Show
   End Select
End Sub
'2013/3/6 end

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

'add by sonia 2021/12/28
Private Sub mnuChUser_Click()
   frmChgUser.Show
End Sub
'end 2021/12/28

'cancel by sonia 2015/5/29 已改全所刷卡,不必再核對打卡卡片
'Private Sub mnu2324_Click(Index As Integer)
'   Select Case Index
'      Case 1 '每日假單簽收明細表
'         If CheckUse("frm180502", strExec) Then
'            frm180502.Show
'         End If
'   End Select
'End Sub
'
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

'Removed by Morgan 2024/8/7
'Private Sub ChkConnection()
'   Dim stSQL As String, intQ As Integer
'   Dim stErr As String
'   stSQL = "select * from dual"
'   intQ = 1
'   ClsLawReadRstMsg intQ, stSQL, , True, stErr
'   If intQ <> 1 Then
'      PUB_WriteLog stErr
'      ConnectToServer_1
'   End If
'End Sub

'add by nickc 2005/05/02
Private Sub Timer3_Timer()
   Static dtNow As Date
   
   'Added by Morgan 2021/8/25 每1分鐘檢查連線是否正常(執行檔傳參數 X 可略過)
   'Modified by Morgan 2021/9/1 改每10分鐘檢查
   'Modified by Morgan 2024/8/7 定時執行一次語法以確保跨網段連線時網路不會被切斷(改用統一寫法)
   'If pub_strCommand <> "X" Then
   '   If Now > dtNow Then
   '      dtNow = DateAdd("n", 10, Now)
   '      ChkConnection
   '   End If
   'End If
On Error Resume Next '若有其他軟體也在使用剪貼簿時會發生521(無法開啟剪貼簿)的錯誤(Ex.Word開啟剪貼簿並擷取畫面)
   
   If Now > dtNow Then
      dtNow = DateAdd("n", cntAutoQueryInterval, Now)
      ClsLawReadRstMsg 1, "select * from dual"
   End If
   'end 2024/8/7
   'end 2021/8/25

'Added by Morgan 2016/4/8
If Not Me.ActiveForm Is Nothing Then
   If LCase(Left(Me.ActiveForm.Name, 6)) = "frmacc" Then
       Toolbar1.Visible = True
       StatusBar1.Visible = True
   Else
       Toolbar1.Visible = False
       StatusBar1.Visible = False
   End If
End If
'end 2016/4/8
   
'add by nickc 2005/05/02 電腦中心的不管
If Pub_StrUserSt03 = "M51" Or Pub_Can_Copy_Pic = True Then Exit Sub
     If Clipboard.GetFormat(2) = True And Clipboard.GetFormat(3) = False And Clipboard.GetFormat(1) = False Then
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

'Added by Lydia 2016/12/19 會計師資料／客戶E-Mail資料查詢
Private Sub Main2_15_Click()
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
   'Add by Amy 2024/01/22
   Case "frm090801_14" '接洽單-對造
         Set GetForm = frm090801_14
   'Add By Sindy 2020/5/29
   Case "frm180301"
         Set GetForm = frm180301
   '2020/5/29 END
   'Add by Amy 2023/09/26 共同查詢用
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
   End Select
End Function

'Added by Lydia 2020/03/26  特別開放Casher的收據作業: 不控制權限
Private Sub Main4_1_1_Click(Index As Integer)
    If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   
   Select Case Index
      Case 1 '收據開立作業
         If PUB_GetLock("Frmacc1120", "", "收據開立作業") = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1120.Show
         Me.MousePointer = vbDefault
      Case 2 '收據開立作業-整批
         If PUB_GetLock("Frmacc1123", "", "收據開立作業-批次") = False Then
            Exit Sub
         End If
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1123.Show
         Me.MousePointer = vbDefault
      Case 3 '收據抬頭修改
         tool8_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1140.Show
         Me.MousePointer = vbDefault
      Case 4 '收據/請款單作廢作業
         tool14_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1130.Show
         Me.MousePointer = vbDefault
      Case 5 '收據抬頭基本資料維護
         tool1_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc11p0.Show
         Me.MousePointer = vbDefault
      'Added by Lydia 2020/03/27
      Case 6 '國內應收待處理作業
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc11r0.Show
         Me.MousePointer = vbDefault
   End Select
End Sub

'Added by Lydia 2020/03/26  特別開放Casher的收據作業: 不控制權限
Private Sub Main4_2_1_Click(Index As Integer)
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   
   Select Case Index
        Case 1   '收文與收據資料檢核查詢
            tool3_enabled
            Me.MousePointer = vbHourglass
            MenuDisabled
            Frmacc1280.Show
            Me.MousePointer = vbDefault
        'Added by Lydia 2020/03/30
        Case 2  '本所案號帳目查詢
            If strFormName <> MsgText(601) Then
               Exit Sub
            End If
            tool3_enabled
            Me.MousePointer = vbHourglass
            MenuDisabled
            Frmacc1240.m_strUserOffice = "1" '傳入北所=>不受所別限制
            Frmacc1240.Show
            Me.MousePointer = vbDefault
   End Select
End Sub

'Added by Lydia 2020/03/26  特別開放Casher的收據作業: 不控制權限
Private Sub Main4_3_1_Click(Index As Integer)
   If strFormName <> MsgText(601) Then
      Exit Sub
   End If
   
   Select Case Index
      Case 1 '收據列印
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1410.ProState = "1" '權限: 1.全所 2.該所 Add By Sindy 2021/5/21
         Frmacc1410.Show
         Me.MousePointer = vbDefault
      Case 2 '補開收據列印
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1420.ProState = "1" '權限: 1.全所 2.該所 Add By Sindy 2021/5/21
         Frmacc1420.Show
         Me.MousePointer = vbDefault
      Case 3 '請款單及發票列印
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1610.ProState = "1" '權限: 1.全所 2.該所 Add By Sindy 2021/5/21
         Frmacc1610.Show
         Me.MousePointer = vbDefault
      Case 4 '補開請款單及發票列印
         tool3_enabled
         Me.MousePointer = vbHourglass
         MenuDisabled
         Frmacc1620.ProState = "1" '權限: 1.全所 2.該所 Add By Sindy 2021/5/21
         Frmacc1620.Show
         Me.MousePointer = vbDefault
   End Select
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
