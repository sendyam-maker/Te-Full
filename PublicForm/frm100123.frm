VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100123 
   BorderStyle     =   1  '單線固定
   Caption         =   "期限資料查詢"
   ClientHeight    =   6120
   ClientLeft      =   50
   ClientTop       =   350
   ClientWidth     =   10080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   10080
   Begin VB.CommandButton cmdOK 
      Caption         =   "管制備註(&A)"
      Height          =   375
      Index           =   3
      Left            =   5010
      Style           =   1  '圖片外觀
      TabIndex        =   54
      Top             =   0
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   " 資料寄信箱(&D)"
      Height          =   375
      Index           =   8
      Left            =   3390
      Style           =   1  '圖片外觀
      TabIndex        =   53
      Top             =   0
      Width           =   1350
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "EMail"
      Enabled         =   0   'False
      Height          =   375
      Index           =   7
      Left            =   4710
      Style           =   1  '圖片外觀
      TabIndex        =   52
      Top             =   480
      Width           =   930
   End
   Begin VB.CheckBox Check7 
      Caption         =   "結案中：已送出電子結案單，程序尚未處理"
      Height          =   180
      Left            =   4740
      TabIndex        =   51
      Top             =   2250
      Width           =   4005
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "搜尋(&Q)"
      Height          =   315
      Left            =   3330
      TabIndex        =   10
      Top             =   1335
      Width           =   765
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdAtt 
      Height          =   1905
      Left            =   720
      TabIndex        =   49
      Top             =   3330
      Visible         =   0   'False
      Width           =   5475
      _ExtentX        =   9648
      _ExtentY        =   3369
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "原路徑檔名|檔名|本所案號|電子檔狀況"
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.CheckBox Check5 
      Caption         =   "含此期間法定期限案件(逾本所期限) )"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   4740
      TabIndex        =   18
      Top             =   1200
      Width           =   3225
   End
   Begin VB.CheckBox Check6 
      Caption         =   "只顯示已函知 )"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   8040
      TabIndex        =   19
      Top             =   1200
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm100123.frx":0000
      Left            =   4470
      List            =   "frm100123.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   47
      Top             =   2482
      Width           =   3195
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0000C0C0&
      Caption         =   "結案"
      Height          =   375
      Index           =   6
      Left            =   8130
      Style           =   1  '圖片外觀
      TabIndex        =   30
      Top             =   480
      Width           =   1000
   End
   Begin VB.CheckBox chkNP 
      Caption         =   "未回覆：主管機關或代理人未回覆處理程序(專業部管制,耗時)"
      Height          =   180
      Left            =   4740
      TabIndex        =   23
      Top             =   2040
      Width           =   5145
   End
   Begin VB.CheckBox Check4 
      Caption         =   "未發文：收文未發文"
      Height          =   180
      Left            =   4740
      TabIndex        =   22
      Top             =   1830
      Width           =   5145
   End
   Begin VB.CheckBox Check3 
      Caption         =   "未通知：主管機關或代理人來函專業部尚未通知智權部"
      Height          =   180
      Left            =   4740
      TabIndex        =   21
      Top             =   1620
      Width           =   5145
   End
   Begin VB.CheckBox Check2 
      Caption         =   "未收款：超過付款週期之未收款"
      Height          =   180
      Left            =   4740
      TabIndex        =   20
      Top             =   1410
      Width           =   5145
   End
   Begin VB.CheckBox Check1 
      Caption         =   "未處理、未函知：未收文且未結案"
      Height          =   180
      Left            =   4740
      TabIndex        =   17
      Top             =   990
      Width           =   3405
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "frm100123.frx":0004
      Left            =   3150
      List            =   "frm100123.frx":0023
      TabIndex        =   7
      Top             =   720
      Width           =   750
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0000C0C0&
      Caption         =   "卷宗區"
      Height          =   375
      Index           =   4
      Left            =   7230
      Style           =   1  '圖片外觀
      TabIndex        =   29
      Top             =   480
      Width           =   870
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0000C0C0&
      Caption         =   "收文"
      Height          =   375
      Index           =   5
      Left            =   6330
      Style           =   1  '圖片外觀
      TabIndex        =   28
      Top             =   480
      Width           =   870
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   1050
      MaxLength       =   3
      TabIndex        =   12
      Top             =   1980
      Width           =   612
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   13
      Top             =   1980
      Width           =   1236
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   2
      Left            =   2940
      MaxLength       =   1
      TabIndex        =   14
      Top             =   1980
      Width           =   276
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   3
      Left            =   3240
      MaxLength       =   2
      TabIndex        =   15
      Top             =   1980
      Width           =   420
   End
   Begin VB.TextBox txtCP10 
      Height          =   300
      Left            =   1050
      MaxLength       =   4
      TabIndex        =   16
      Top             =   2310
      Width           =   612
   End
   Begin VB.TextBox systemkind 
      Height          =   300
      Left            =   1050
      TabIndex        =   11
      Text            =   "ALL"
      Top             =   1650
      Width           =   2130
   End
   Begin VB.TextBox txtCU2 
      Height          =   300
      Left            =   2280
      MaxLength       =   9
      TabIndex        =   9
      Top             =   1020
      Width           =   1095
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   300
      Index           =   1
      Left            =   2100
      MaxLength       =   7
      TabIndex        =   6
      Top             =   720
      Width           =   915
   End
   Begin VB.TextBox txtCU1 
      Height          =   300
      Left            =   1050
      MaxLength       =   9
      TabIndex        =   8
      Top             =   1020
      Width           =   1095
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   300
      Index           =   0
      Left            =   1050
      MaxLength       =   7
      TabIndex        =   5
      Top             =   720
      Width           =   915
   End
   Begin VB.TextBox txtSales 
      Height          =   300
      Left            =   1050
      MaxLength       =   6
      TabIndex        =   4
      Top             =   360
      Width           =   915
   End
   Begin VB.TextBox txtSalesArea 
      Height          =   300
      Left            =   1050
      TabIndex        =   1
      Top             =   30
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H008080FF&
      Caption         =   "內商延展　結案通知(&M)"
      Height          =   405
      Index           =   2
      Left            =   8400
      Style           =   1  '圖片外觀
      TabIndex        =   31
      Top             =   2430
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件進度(&C)"
      Height          =   375
      Index           =   1
      Left            =   7290
      Style           =   1  '圖片外觀
      TabIndex        =   25
      Top             =   0
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "基本資料(&B)"
      Height          =   375
      Index           =   0
      Left            =   6150
      Style           =   1  '圖片外觀
      TabIndex        =   24
      Top             =   0
      Width           =   1110
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   9210
      Style           =   1  '圖片外觀
      TabIndex        =   27
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   375
      Left            =   8445
      Style           =   1  '圖片外觀
      TabIndex        =   26
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtSalesArea1 
      Height          =   300
      Left            =   2040
      TabIndex        =   2
      Top             =   30
      Width           =   915
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3195
      Left            =   60
      TabIndex        =   32
      Top             =   2850
      Width           =   9915
      _ExtentX        =   17480
      _ExtentY        =   5627
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   7
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.TextBox txtZone 
      Height          =   285
      Left            =   4455
      MaxLength       =   1
      TabIndex        =   0
      Top             =   2490
      Visible         =   0   'False
      Width           =   915
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7620
      Top             =   2490
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSForms.ComboBox Combo3 
      Height          =   330
      Left            =   1050
      TabIndex        =   3
      Top             =   360
      Width           =   1920
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3387;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCuName 
      Height          =   300
      Left            =   1380
      TabIndex        =   55
      Top             =   1335
      Width           =   1875
      VariousPropertyBits=   679495707
      MaxLength       =   40
      Size            =   "3307;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCuNam 
      Caption         =   "客戶中文名稱："
      Height          =   180
      Left            =   120
      TabIndex        =   50
      Top             =   1395
      Width           =   1275
   End
   Begin VB.Label Label6 
      Caption         =   "("
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   4650
      TabIndex        =   48
      Top             =   1200
      Width           =   75
   End
   Begin VB.Label Label7 
      Caption         =   "("
      Height          =   255
      Left            =   3060
      TabIndex        =   46
      Top             =   743
      Width           =   75
   End
   Begin VB.Label LblCntTime 
      AutoSize        =   -1  'True
      Caption         =   "執行時間："
      Height          =   180
      Left            =   120
      TabIndex        =   45
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "個月)"
      Height          =   195
      Index           =   1
      Left            =   3900
      TabIndex        =   44
      Top             =   773
      Width           =   435
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   120
      TabIndex        =   43
      Top             =   2025
      Width           =   900
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Left            =   120
      TabIndex        =   42
      Top             =   2355
      Width           =   900
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Left            =   120
      TabIndex        =   41
      Top             =   1695
      Width           =   900
   End
   Begin VB.Line Line3 
      X1              =   2070
      X2              =   2340
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Left            =   120
      TabIndex        =   40
      Top             =   1065
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "業務區："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   39
      Top             =   90
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所期限："
      Height          =   180
      Left            =   120
      TabIndex        =   38
      Top             =   765
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "所別："
      Height          =   180
      Left            =   3750
      TabIndex        =   37
      Top             =   2542
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Line Line1 
      X1              =   1920
      X2              =   2190
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   120
      TabIndex        =   36
      Top             =   435
      Width           =   900
   End
   Begin MSForms.Label lblSalesName 
      Height          =   180
      Left            =   2040
      TabIndex        =   35
      Top             =   400
      Width           =   2320
      VariousPropertyBits=   268435483
      Caption         =   "lblSalesName"
      Size            =   "4092;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblZone 
      AutoSize        =   -1  'True
      Caption         =   "（1北所 2中所 3南所 4高所）"
      Height          =   180
      Left            =   5400
      TabIndex        =   34
      Top             =   2415
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   105
      TabIndex        =   33
      Top             =   5430
      Width           =   45
   End
   Begin VB.Line Line2 
      X1              =   1830
      X2              =   2100
      Y1              =   150
      Y2              =   150
   End
End
Attribute VB_Name = "frm100123"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2014/6/13 調整SQL的抓法,使得讀取資料速度快些
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'Memo by Lydia 2019/07/01 表單名稱:智權人員期限資料查詢 =>期限資料查詢
'2005/7/5整理
'Memo by Lydia 2021/05/10 Form2.0已修改txtCuName、Combo3、lblSalesName、grdDataList改字型=新細明體-ExtB
Option Explicit

Dim bolShowMsgBox As Boolean ', bolSelData As Boolean
'2005/08/18 nick 紀錄作用按鍵
Public cmdState As Integer
Dim i As Integer, j As Integer
Dim StrTag As String
'add by nickc 2008/01/18
Dim stST05 As String ', stST15 As String
'Add by Amy 2014/05/15
Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序 Add By Sindy 2014/6/19
'add by nickc 2007/03/29
Dim StrTmpCp09 As String
'Dim TmpArr As Variant
Dim StrTmpNp22 As String
'Dim TmpArrNp22 As Variant
'add by nickc 2005/09/20
Dim StrTmpCp01020304 As String
Dim StrCompCp01020304 As Variant 'Add By Sindy 2015/3/11
Dim TmpArrCaseNo As Variant 'Add By Sindy 2015/3/11
'Added by Lydia 2015/06/18 +客戶中文名稱查詢
Public m_strCustCode As String
Public m_blnOneRec As Boolean
Dim m_AttachPath As String 'Added by Morgan 2016/3/1
Dim m_strListPer As String 'Add By Sindy 2016/5/4
Dim arrID                  'add by sonia 2016/6/7
Dim strMCTF As String 'Add by Amy 2017/01/12
'Add by Amy 2020/03/25
Const NotSOutXls As String = "71011,69008,68005" '非S部門可產生Excel人員
Dim arrGridHeadText '從SetDataListWidth搬出來
'Dim bolAreaMan As Boolean '下拉選單有區主管
Dim xlsFileName As String
'Added by Lydia 2021/05/10 Grid的欄位位置
Dim colPKey As Integer 'PKey，原位置21
Dim colType As Integer '事件，原位置5
Dim colCaseNo As Integer '本所案號，原位置9
Dim colCP06 As Integer '本所期限，原位置7
Dim colCPM As Integer '案件性質，原位置12
Dim colNP01 As Integer '收文號，原位置19
Dim colNP22 As Integer '序號，原位置20
Dim colCMR04 As Integer 'Added by Lydia 2021/05/28 管制備註欄位
'end 2021/05/10
Const m_NewStartDate As String = "20210604" 'Added by Lydia 2021/05/10 (啟用日)增加管制備註功能：調整版面


Private Sub SetDataListWidth()
Dim arrGridHeadWidth
Dim iRow As Integer

   'Added by Lydia 2021/05/10 (啟用日)增加管制備註功能：調整版面
   If strSrvDate(1) >= m_NewStartDate Then
       Call SetDataListWidth_1
       Exit Sub
   End If
   'end 2021/05/10
   
   'Modified by Morgan 2012/8/21 新增"約定期限","部門"改不顯示
   'edit by nickc 2005/09/06
   'arrGridHeadText = Array("V", "所別", "部門", "管制人", "智權人員", "事件" _
   '                  , "本所期限", "法定期限", "本所案號", "案件名稱" _
   '                  , "案件性質", "承辦人", "收文日", "發文日", "申請人" _
   '                  , "申請國家", "申請案號")
   '
   'arrGridHeadWidth = Array(200, 480, 680, 680, 980, 680 _
   '                     , 750, 750, 1200, 2000 _
   '                     , 1000, 800, 700, 700, 2000 _
   '                     , 800, 1200)
   'arrGridHeadText = Array("V", "所別", "部門", "管制人", "智權人員", "事件" _
   '                  , "本所期限", "法定期限", "本所案號", "案件名稱" _
   '                  , "案件性質", "承辦人", "收文日", "發文日", "申請人" _
   '                  , "申請國家", "申請案號", "收文號")
   '
   'arrGridHeadWidth = Array(200, 480, 680, 680, 680, 680 _
   '                     , 820, 750, 1400, 2000 _
   '                     , 1000, 800, 700, 700, 2000 _
   '                     , 800, 1200, 950)
   'edit by nickc 2005/09/29 M51及分所加入分所號在本所案號後
   'edit by nickc  2007/03/28 加入序號
   'edit by Sindy  2010/01/15 加入PKey
   'arrGridHeadText = Array("V", "所別", "部門", "管制人", "智權人員", "事件" _
                     , "本所期限", "法定期限", "本所案號", "分所號", "案件名稱" _
                     , "案件性質", "承辦人", "收文日", "發文日", "申請人" _
                     , "申請國家", "申請案號", "收文號")
   arrGridHeadText = Array("V", "所別", "部門", "管制人", "智權人員", "事件", "約定期限" _
                     , "本所期限", "法定期限", "本所案號", "分所號", "案件名稱" _
                     , "案件性質", "承辦人", "收文日", "發文日", "申請人" _
                     , "申請國家", "申請案號", "收文號", "序號", "PKey")

Dim iDep As String
   
   iDep = PUB_GetST06(strUserNum)
   'edit by nickc 2005/09/29 M51所別及分所號都顯示,分所不顯示所別但加入分所號在本所案號後,北所非M51顯示所別不顯示分所號
   '2010/8/24 modify by sonia 調整日期欄位寬度
   'Modified by Morgan 2011/11/22 考慮拆收據情形,收文號欄位加大(未收款資料的收據號放在該欄位)
   If GetStaffDepartment(strUserNum) = "M51" Then
      'edit by nickc  2007/03/28 加入序號
      'edit by Sindy  2010/01/15 加入PKey
      'arrGridHeadWidth = Array(200, 250, 680, 680, 680, 680 _
                           , 820, 750, 1200, 670, 2000 _
                           , 1000, 800, 700, 700, 2000 _
                           , 800, 1200, 950)
      arrGridHeadWidth = Array(200, 250, 0, 680, 680, 680, 850 _
                           , 900, 850, 1200, 670, 1700 _
                           , 1000, 800, 850, 850, 2000 _
                           , 800, 1200, 1800, 0, 0)
   Else
      If iDep = "1" Then
         'edit by nickc  2007/03/28 加入序號
         'edit by Sindy  2010/01/15 加入PKey
         'arrGridHeadWidth = Array(200, 480, 680, 680, 680, 680 _
                              , 820, 750, 1200, 0, 2000 _
                              , 1000, 800, 700, 700, 2000 _
                              , 800, 1200, 950)
         arrGridHeadWidth = Array(200, 480, 0, 680, 680, 680, 850 _
                              , 900, 850, 1200, 0, 1920 _
                              , 1000, 800, 850, 850, 2000 _
                              , 800, 1200, 1800, 0, 0)
      Else
         'edit by nickc  2007/03/28 加入序號
         'edit by Sindy  2010/01/15 加入PKey
         'arrGridHeadWidth = Array(200, 0, 680, 680, 680, 680 _
                              , 820, 750, 1200, 670, 2000 _
                              , 1000, 800, 700, 700, 2000 _
                              , 800, 1200, 950)
         arrGridHeadWidth = Array(200, 0, 0, 680, 680, 680, 850 _
                              , 900, 850, 1200, 670, 1920 _
                              , 1000, 800, 850, 850, 2000 _
                              , 800, 1200, 1800, 0, 0)
      End If
   End If
   'If Me.txtZone.Locked = True Or Me.txtZone.Enabled = False Then
   '   arrGridHeadWidth(1) = 0
   'End If
   'If (Me.txtSalesArea.Locked = True And Me.txtSalesArea1.Locked = True) Or (Me.txtSalesArea.Enabled = False And Me.txtSalesArea1.Enabled = False) Then
   '   arrGridHeadWidth(2) = 0
   'End If
   '2007/8/27 cancel by sonia 因T-055520,96/8/29客戶轉莊宏宇但最後收文智權人員為王文德,若不看管制人會很怪
   'If Me.txtSales.Locked = True Or Me.txtSales.Enabled = False Then
   '   arrGridHeadWidth(3) = 0
   'End If
   '2007/8/27 end
   'grdDataList.MergeCol(1) = True
   'grdDataList.MergeCol(2) = True
   'grdDataList.MergeCol(3) = True
   'grdDataList.MergeCol(4) = True
   'grdDataList.MergeCol(5) = True
   'grdDataList.MergeCol(6) = True
   grdDataList.MergeCells = flexMergeRestrictColumns
   grdDataList.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To grdDataList.Cols - 1
      grdDataList.row = 0
      grdDataList.col = iRow
      grdDataList.Text = arrGridHeadText(iRow)
      grdDataList.ColWidth(iRow) = arrGridHeadWidth(iRow)
      grdDataList.CellAlignment = flexAlignLeftCenter
   Next
   
   'Added by Lydia 2021/05/10 取得Grid的欄位位置
   If colPKey = 0 Then
       colPKey = PUB_MGridGetId("PKey", Me.grdDataList)
       colType = PUB_MGridGetId("事件", Me.grdDataList)
       colCaseNo = PUB_MGridGetId("本所案號", Me.grdDataList)
       colCP06 = PUB_MGridGetId("本所期限", Me.grdDataList)
       colCPM = PUB_MGridGetId("案件性質", Me.grdDataList)
       colNP01 = PUB_MGridGetId("收文號", Me.grdDataList)
       colNP22 = PUB_MGridGetId("序號", Me.grdDataList)
   End If
   'end 2021/05/10
End Sub

Private Function ConstrainCheck() As Boolean
Dim bolCancel As Boolean
Dim intErrCol As Integer
   
   ConstrainCheck = True
   
   'Modify By Sindy 2010/6/17
   If Text1(0) = "" Or Text1(1) = "" Then
      If txtCloseDate(0) = "" Then
         MsgBox "請輸入本所期限起日！", vbExclamation
         txtCloseDate(0).SetFocus
         txtCloseDate_GotFocus (0)
         ConstrainCheck = False
         Exit Function
      Else
         bolCancel = False
         Call txtCloseDate_Validate(0, bolCancel)
         If bolCancel = True Then
            ConstrainCheck = False
            Exit Function
         End If
      End If
      If txtCloseDate(1) = "" Then
         MsgBox "請輸入本所期限迄日！", vbExclamation
         txtCloseDate(1).SetFocus
         txtCloseDate_GotFocus (1)
         ConstrainCheck = False
         Exit Function
      Else
         bolCancel = False
         Call txtCloseDate_Validate(1, bolCancel)
         If bolCancel = True Then
            ConstrainCheck = False
            Exit Function
         End If
      End If
   Else
      If Text1(2) = "" Then Text1(2) = "0"
      If Text1(3) = "" Then Text1(3) = "00"
   End If
   
   'Add by Amy 2020/03/25 +有下拉選單
   If Combo3.Visible = True Then
      Call Combo3_LostFocus 'Add By Sindy 2020/7/15 讓人員按Enter,須再啟動此函數,txtSales欄位值才會置換
      If Combo3 = MsgText(601) Then
          Call Combo3_Validate(bolCancel)
          If bolCancel = True Then
              Combo3.SetFocus
              ConstrainCheck = False
              Exit Function
          End If
      ElseIf txtSales = MsgText(601) Then
          txtSales = Mid(Combo3, 1, Val(InStr(Combo3, " ")) - 1)
      End If
   End If
   
   'Add By Sindy 2009/05/14
   Call txtSales_Validate(bolCancel)
   If bolCancel = True Then
      'Modify by Amy 2020/03/25 +有下拉選單
      If Combo3.Visible = True Then
         Combo3.SetFocus
      'Modified by Lydia 2021/05/20 排除隱藏
      'ElseIf txtSales.Enabled = True Then
      ElseIf txtSales.Enabled = True And txtSales.Visible = True Then
         txtSales.SetFocus
         txtSales_GotFocus
      End If
      ConstrainCheck = False
      Exit Function
   End If
   
   'Modify By Sindy 2020/7/29 檢查部門欄位
   'Modify By Sindy 2025/8/11 +, txtZone
   If PUB_ChkFormSalesDept(strUserNum, txtSales, txtSalesArea, txtSalesArea1, intErrCol, txtZone) = False Then
      If intErrCol = 0 Then
         If txtSales.Visible = True Then 'Added by Lydia 2021/05/20 排除隱藏
            txtSales.SetFocus
            txtSales_GotFocus
         End If  'Added by Lydia 2021/05/20
      ElseIf intErrCol = 1 Then
         txtSalesArea.SetFocus
         txtSalesArea_GotFocus
      Else
         txtSalesArea1.SetFocus
         txtSalesArea1_GotFocus
      End If
      ConstrainCheck = False
      Exit Function
   End If
   
   'add by nickc 2007/01/19
   If Trim(txtCU1) <> "" Or Trim(txtCU2) <> "" Then
      If Mid(txtCU1, 1, 6) <> Mid(txtCU2, 1, 6) Then
          MsgBox "申請人前6碼必須相同！", vbExclamation
          txtCU1.SetFocus
          txtCU1_GotFocus
          ConstrainCheck = False
          Exit Function
      End If
   End If
    
    'Add By Sindy 2014/6/17
    'Modify By Sindy 2015/9/17 +Check7
    If Check1.Value = 0 And Check2.Value = 0 And Check3.Value = 0 And _
       Check4.Value = 0 And chkNP.Value = 0 And Check7.Value = 0 Then
      MsgBox "查詢類別至少要勾選一項！", vbExclamation
      ConstrainCheck = False
      Check1.SetFocus
      Exit Function
    End If
    '2014/6/17 END
End Function

Public Function doQuery() As Boolean
Dim stCon As String, stConST As String
Dim stCon1 As String, stCon2 As String, stCon3 As String, stCon4 As String
Dim stCon5 As String, stCon6 As String, strCP13 As String 'Add By Sindy 2011/6/21
'Dim stCon7 As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strInData As String
'Dim stConAcc0k0 As String    '2009/11/19 add by sonia 調整未收款語法
Dim stVTBw As String 'Add by Morgan 2010/1/27 預定收款日資料
Dim stIdList As String, stConId As String 'Add by Morgan 2010/1/29
Dim bolChgSystemkind As Boolean, strOldSystemkind As String
Dim stCon8 As String 'Add by Morgan 2011/8/15
Dim stCon9 As String 'Add by Morgan 2012/8/21
Dim strChkSql As String 'Add By Sindy 2014/6/12
   
'Added by Lydia 2021/05/10 (啟用日)增加管制備註功能：調整版面
If strSrvDate(1) >= m_NewStartDate Then
    Call doQuery_1
    Exit Function
End If
'end 2021/05/10

'Memo by Lydia 2021/05/10
'相關Table:
'R100123: 確定刪除Table
'R100123_1：期限資料查詢(frm100123) 增加管制備註功能－調整後使用
'R100123_2: 期限資料查詢 (frm100123) - 調整前使用
'end 2021/05/10

   m_blnColOrderAsc = True 'Add By Sindy 2014/6/19
   LblCntTime.Caption = "執行時間：" & Format(ServerTime, "##:##:##") 'Add By Sindy 2014/6/12
   'add by nickc 2007/01/19
   Dim stConP As String, stConT As String, stConS As String, stConL As String, stConH As String
   stConP = "": stConT = "": stConS = "": stConL = "": stConH = ""
   
   'add by nickc 2007/01/23
   Dim stCon1_1 As String, stCon2_1 As String
   stCon1_1 = "": stCon2_1 = ""
   
   'add by nickc 2008/04/24 加入未收款
   stCon4 = ""
   'Add By Sindy 2011/6/21
   stCon5 = "" '未續簽
   stCon6 = "" '未回執
'   stCon7 = ""
   
   stCon8 = "" 'Add by Morgan 2011/8/15
   
   stCon = "": stConST = "": stCon1 = "": stCon2 = "": stCon3 = ""
'   stConAcc0k0 = ""  '2009/11/19 add by sonia
   
   'Modify By Sindy 2010/6/17
   bolChgSystemkind = False
   If Text1(0) = "" Or Text1(1) = "" Then
      If systemkind = "" Then
          systemkind = "ALL"
      End If
   Else
      bolChgSystemkind = True
      strOldSystemkind = systemkind
      systemkind = Text1(0).Text
   End If
   
   '所別
'edit by nickc 2005/09/06 不檢查所別了
'   If txtZone <> "" Then
'      stConST = stConST & " and s2.st06 = '" & txtZone & "'"
'   End If
   '2005/9/8 ADD BY SONIA 蔣律師要控制所別
'cancel by sonia 2014/6/9
'   If strUserNum = "79037" Then
'      stConST = stConST & " and s2.st06 = '" & pub_strUserOffice & "'"
'      'Add By Sindy 2011/6/21 未續簽&未回執
'      'stCon7 = stCon7 & " and s2.st06 = '" & pub_strUserOffice & "'"
'      stCon5 = stCon5 & " and s2.st06 = '" & pub_strUserOffice & "'"
'      stCon6 = stCon6 & " and s2.st06 = '" & pub_strUserOffice & "'"
'   End If
'end 2014/6/9
   '2005/9/8 END
   'add by sonia 2016/6/7 林柄佑要控制所別
   If strUserNum = "82026" Then
      stConST = stConST & " and s2.st06 = '" & pub_strUserOffice & "'"
      'stCon7 = stCon7 & " and s2.st06 = '" & pub_strUserOffice & "'"
      stCon5 = stCon5 & " and s2.st06 = '" & pub_strUserOffice & "'"
      stCon6 = stCon6 & " and s2.st06 = '" & pub_strUserOffice & "'"
   End If
   'end 2016/6/7
   
   '2005/9/12 ADD BY SONIA 陳經理查詢所有智權人員要控制系統類別
   If strUserNum = "68005" And txtSales <> "68005" Then
      systemkind = "CFT,FCT,S,CFC"
   End If
   '2005/9/12 END
   
   '區別
   'Add by Amy 2014/05/15
    'Modify by Amy 2019/02/12 總經理業務工作代理人員
   If bolSpecMan = True And (InStr(strSpecCode, "A8") > 0 Or InStr(strSpecCode, "總經理業務工作代理人員") > 0) And txtSales <> strUserNum Then
            '開放專利處部份智權同仁資料給彥葶代為處理,不考慮業務區(因彥葶與開放的智權同仁業務區不同)
   'end 2014/05/15
   'Modify By Sindy 98/03/11 若智權人員為80030時, 不限制區別
   'Modify By Sindy 2009/05/12
   '若為帶人主管權限時,查詢之智權人員編號非本人時,不限制區別
   '2009/12/16 MODIFY BY SONIA 加巨京專利給郭雅娟79075看,所以不限制區別
   'modify by sonia 2016/6/7 帶人主管條件改寫法
   'If txtSales = "80030" Or txtSales = "79075" Or _
      (Trim(txtSales) <> "" And PUB_GetST05(strUserNum) = "SA" And txtSales.Enabled = True And txtSales <> strUserNum) Then
   ElseIf txtSales = "80030" Or txtSales = "79075" Or _
      (Trim(txtSales) <> "" And PUB_GetST05Limits(strUserNum) = True And txtSales.Enabled = True And txtSales <> strUserNum) Then
      '不限制區別
  
   'add by sonia 2016/6/7 查自己資料不限制區別,因為有調區問題
   ElseIf txtSales = strUserNum Then
   'end 2016/6/7
   Else
         If txtSalesArea <> "" Then
'            If chkNP.Value = 1 Then
'                stCon1 = stCon1 & " and cp12||'' >= '" & txtSalesArea & "'"
'            Else
                stCon1 = stCon1 & " and s1.st15>='" & txtSalesArea & "'"
'            End If
            stCon2 = stCon2 & " and cp12||''>='" & txtSalesArea & "'"
            stCon3 = stCon3 & " and s2.st15>='" & txtSalesArea & "'"
            'add by nickc 2008/04/24 加入未收款
            stCon4 = stCon4 & " and a0k22||''>='" & txtSalesArea & "'"
'            stConAcc0k0 = stConAcc0k0 & " and a0k22||''>='" & txtSalesArea & "'"  '2009/11/19 add by sonia
            'Add By Sindy 2011/6/21 未續簽&未回執
'            'stCon7 = stCon7 & " and rcp12||''>='" & txtSalesArea & "'"
            stCon5 = stCon5 & " and cp12||''>='" & txtSalesArea & "'"
            stCon6 = stCon6 & " and cp12||''>='" & txtSalesArea & "'"
         End If
         If txtSalesArea1 <> "" Then
'            If chkNP.Value = 1 Then
'                stCon1 = stCon1 & " and cp12||'' <= '" & txtSalesArea1 & "'"
'            Else
                stCon1 = stCon1 & " and s1.st15<='" & txtSalesArea1 & "'"
'            End If
            stCon2 = stCon2 & " and cp12||''<='" & txtSalesArea1 & "'"
            stCon3 = stCon3 & " and s2.st15<='" & txtSalesArea1 & "'"
            'add by nickc 2008/04/24 加入未收款
            stCon4 = stCon4 & " and a0k22||''<='" & txtSalesArea1 & "'"
'            stConAcc0k0 = stConAcc0k0 & " and a0k22||''<='" & txtSalesArea1 & "'"  '2009/11/19 add by sonia
            'Add By Sindy 2011/6/21 未續簽&未回執
'            'stCon7 = stCon7 & " and rcp12||''<='" & txtSalesArea1 & "'"
            stCon5 = stCon5 & " and cp12||''<='" & txtSalesArea1 & "'"
            stCon6 = stCon6 & " and cp12||''<='" & txtSalesArea1 & "'"
         End If
   End If
'2005/6/24 end
   
   '智權人員
   If Trim(txtSales) <> "" Then
        'add by nickc 2006/06/23 加入區主管若是輸入自己的編號時，要看見 自己的 + 離職智權人員 + 虛建智權人員的資料
        '2006/11/29 MODIFY BY SONIA 加入69005之控制
        'If (strUserNum = "74018" And txtSales = "74018") Or (strUserNum = "78007" And txtSales = "78007") Or (strUserNum = "71011" And txtSales = "71011") Or (strUserNum = "67002" And txtSales = "67002") Or (PUB_GetST05(strUserNum) = "SM" And strUserNum <> "71003") Or (strUserNum = "71003" And txtSales = "71003" And txtSalesArea = "S23" And txtSalesArea1 = "S23") Then
        'edit by nickc 2008/04/24
        'If (strUserNum = "74018" And txtSales = "74018") Or (strUserNum = "78007" And txtSales = "78007") Or (strUserNum = "71011" And txtSales = "71011") Or (strUserNum = "67002" And txtSales = "67002") Or (PUB_GetST05(strUserNum) = "SM" And strUserNum <> "71003") Or (strUserNum = "71003" And txtSales = "71003" And txtSalesArea = "S23" And txtSalesArea1 = "S23") Or (PUB_GetST05(strUserNum) = "SM" And strUserNum <> "69005") Or (strUserNum = "69005" And txtSales = "69005" And txtSalesArea = "S15" And txtSalesArea1 = "S15") Then

'edit by nickc 2008/04/25 改共用
''Modify By Sindy 98/02/27
        'If (strUserNum = "74018" And txtSales = "74018") Or (strUserNum = "78007" And txtSales = "78007") Or (strUserNum = "71011" And txtSales = "71011") Or (strUserNum = "67002" And txtSales = "67002") Or (PUB_GetST05(strUserNum) = "SM" And strUserNum <> "71003" And strUserNum <> "69005") Or (strUserNum = "71003" And txtSales = "71003" And txtSalesArea = "S23" And txtSalesArea1 = "S23") Or (strUserNum = "69005" And txtSales = "69005" And txtSalesArea = "S15" And txtSalesArea1 = "S15") Then
        If (strUserNum <> "80030" And txtSales <> "80030") Then
''98/02/27 End
            
        'Add by Amy 2014/05/15 +if
        'Modify by Amy 2019/02/12 總經理業務工作代理人員,可處理總經理員工編號
        If bolSpecMan = True And (InStr(strSpecCode, "A8") > 0 Or InStr(strSpecCode, "總經理業務工作代理人員") > 0) And txtSales <> strUserNum Then
            '開放專利處部份智權同仁資料給彥葶代為處理,不考慮業務區(因彥葶與開放的智權同仁業務區不同)
            stIdList = PUB_GetSalesList(Trim(txtSales))
        'Add by Amy 2017/01/12 +MCTF控制資料(屬於MCTF特殊人員查個人時也查出MCTF字頭的期限資料)
        'Remove by Lydia 2017/07/21 併入PUB_GetSalesList
        'ElseIf Left(strMCTF, 4) = "MCTF" Then
        '    stIdList = "'" & txtSales & "','" & strMCTF & "'"
        'end 2017/07/21
        Else
            'Add by Morgan 2010/1/29 若不是多員工編號時用 = 算符來加速查詢
            stIdList = PUB_GetSalesList(Trim(txtSales), txtSalesArea, txtSalesArea1, txtZone)
        End If
        'end 2014/05/15
        
            If InStr(stIdList, ",") = 0 Then
               stConId = " = " & stIdList & " "
            Else
               stConId = " in (" & stIdList & " ) "
            End If
            'end 2010/1/29
            
            '2010/5/10 add by sonia 因中所有跨區帶人故離職智權人員的帶人主管不考慮業務區條件
            If Pub_StrST52 Then
               stCon1 = "": stCon2 = "": stCon3 = "": stCon4 = "" ': stConAcc0k0 = ""
            End If
            '2010/5/10 end
            
'            If chkNP.Value = 1 Then
'               'edit by nickc 2008/04/25 改共用
'               'stCon1 = stCon1 & " and cp13||'' in (" & GetNotInOfficeAndFalseStaff(txtSalesArea, txtSalesArea1) & "'" & txtSales & "' ) "
'               'Modify by Morgan 2010/1/29
'               'stCon1 = stCon1 & " and cp13||'' in (" & PUB_GetSalesList(txtSales, txtSalesArea, txtSalesArea1, pub_strUserOffice) & " ) "
'               stCon1 = stCon1 & " and cp13||'' " & stConId
'            Else
               'edit by nickc 2008/04/25 改共用
               'stCon1 = stCon1 & " and np10 in (" & GetNotInOfficeAndFalseStaff(txtSalesArea, txtSalesArea1) & "'" & txtSales & "' ) "
               'Modify by Morgan 2010/1/29
               'stCon1 = stCon1 & " and np10 in (" & PUB_GetSalesList(txtSales, txtSalesArea, txtSalesArea1, pub_strUserOffice) & " ) "
               stCon1 = stCon1 & " and np10 " & stConId
               'end 2010/1/29
'            End If
            'edit by nickc 2008/04/25 改共用
            'stCon2 = stCon2 & " and cp13 in (" & GetNotInOfficeAndFalseStaff(txtSalesArea, txtSalesArea1) & "'" & txtSales & "' ) "
            'stCon3 = stCon3 & " and ss01 in (" & GetNotInOfficeAndFalseStaff(txtSalesArea, txtSalesArea1) & "'" & txtSales & "' ) "
            'Modify by Morgan 2010/1/29
            'stCon2 = stCon2 & " and cp13 in (" & PUB_GetSalesList(txtSales, txtSalesArea, txtSalesArea1, pub_strUserOffice) & " ) "
            'stCon3 = stCon3 & " and ss01 in (" & PUB_GetSalesList(txtSales, txtSalesArea, txtSalesArea1, pub_strUserOffice) & " ) "
            stCon2 = stCon2 & " and cp13 " & stConId
            stCon3 = stCon3 & " and ss01 " & stConId
            'end 2010/1/29
            'add by nickc 2008/04/24 加入未收款
            'Modify by Morgan 2010/1/29
            'stCon4 = stCon4 & " and a0k20||'' in (" & PUB_GetSalesList(txtSales, txtSalesArea, txtSalesArea1, pub_strUserOffice) & " ) "
            'stConAcc0k0 = stConAcc0k0 & " and a0k20||'' in (" & PUB_GetSalesList(txtSales, txtSalesArea, txtSalesArea1, pub_strUserOffice) & " ) "  '2009/11/19 add by sonia
            stCon4 = stCon4 & " and a0k20||'' " & stConId
'            stConAcc0k0 = stConAcc0k0 & " and a0k20||'' " & stConId
            'end 2010/1/29
            'Add By Sindy 2011/6/21 未續簽&未回執
'            'stCon7 = stCon7 & " and rcp13 " & stConId
            stCon5 = stCon5 & " and cp13 " & stConId
            stCon6 = stCon6 & " and cp13 " & stConId
        Else
           'Modify By Sindy 98/02/27 查80030洪琬姿時同時查F4103
            If txtSales = "80030" Then

               StrSQLa = "select ST01 from STAFF where ST04<>'1' and ST03 like 'F1%' "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               strInData = "'80030','F4103'"
               If rsA.RecordCount > 0 Then
                  rsA.MoveFirst
                  Do While rsA.EOF = False
                     strInData = strInData & ",'" & rsA.Fields(0).Value & "'"
                     rsA.MoveNext
                  Loop
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing

'               If chkNP.Value = 1 Then
'                  stCon1 = stCon1 & " and cp13||'' IN (" & strInData & ") "
'               Else
                  stCon1 = stCon1 & " and np10 IN (" & strInData & ") "
'               End If
               stCon2 = stCon2 & " and cp13='" & Trim(txtSales) & "'"
               stCon3 = stCon3 & " and ss01='" & Trim(txtSales) & "'"
               'add by nickc 2008/04/24 加入未收款
               stCon4 = stCon4 & " and a0k20||'' IN (" & strInData & ") "
'               stConAcc0k0 = stConAcc0k0 & " and a0k20||'' IN (" & strInData & ") "  '2009/11/19 add by sonia
               'Add By Sindy 2011/6/21 未續簽&未回執
'               'stCon7 = stCon7 & " and rcp13='" & txtSales & "'"
               stCon5 = stCon5 & " and cp13='" & Trim(txtSales) & "'"
               stCon6 = stCon6 & " and cp13='" & Trim(txtSales) & "'"
            Else
            '98/02/27 END
'               If chkNP.Value = 1 Then
'                  stCon1 = stCon1 & " and cp13||'' = '" & txtSales & "'"
'               Else
                  stCon1 = stCon1 & " and np10='" & Trim(txtSales) & "' "
'               End If
               stCon2 = stCon2 & " and cp13='" & Trim(txtSales) & "'"
               stCon3 = stCon3 & " and ss01='" & Trim(txtSales) & "'"
               'add by nickc 2008/04/24 加入未收款
               stCon4 = stCon4 & " and a0k20||''='" & Trim(txtSales) & "'"
'               stConAcc0k0 = stConAcc0k0 & " and a0k20||''='" & txtSales & "'"  '2009/11/19 add by sonia
               'Add By Sindy 2011/6/21 未續簽&未回執
'               'stCon7 = stCon7 & " and rcp13='" & txtSales & "'"
               stCon5 = stCon5 & " and cp13='" & Trim(txtSales) & "'"
               stCon6 = stCon6 & " and cp13='" & Trim(txtSales) & "'"
            End If
        End If
        '98/02/27 END
   'Modify by Amy 2014/05/15
   '智權人員 為空
   Else
        If bolSpecMan = True And InStr(strSpecCode, "A8") > 0 Then
            'A2023彥葶登入,未輸智權人員-設定查A7人員
            stConId = " in ('" & Replace(Pub_GetSpecMan("A7"), ";", "','") & "')"
'            If chkNP.Value = 1 Then
'               stCon1 = stCon1 & " and cp13||'' " & stConId
'            Else
               stCon1 = stCon1 & " and np10 " & stConId
'            End If
            stCon2 = stCon2 & " and cp13 " & stConId
            stCon3 = stCon3 & " and ss01 " & stConId
            '加入未收款
            stCon4 = stCon4 & " and a0k20||'' " & stConId
'            stConAcc0k0 = stConAcc0k0 & " and a0k20||'' " & stConId
            '未續簽&未回執
'            'stCon7 = stCon7 & " and rcp13 " & stConId
            stCon5 = stCon5 & " and cp13 " & stConId
            stCon6 = stCon6 & " and cp13 " & stConId
        End If
   'end 2014/05/15
   End If
   
   stCon8 = stCon2 'Add by Morgan 2011/8/15 未收款(新)
   
   'add by nickc 2007/01/23
   stCon1_1 = stCon1
   stCon2_1 = stCon2
   stCon9 = stCon1 'Added by Morgan 2012/8/21
   
   '本所期限
   If txtCloseDate(0) <> "" Then
      If Check5.Value = 0 Then 'Add By Sindy 2014/6/24 +if
         stCon1 = stCon1 & " and np08>=" & ChangeTStringToWString(txtCloseDate(0))
      End If
      stCon9 = stCon9 & " and np23>=" & ChangeTStringToWString(txtCloseDate(0)) 'Added by Morgan 2012/8/21
      
      stCon2 = stCon2 & " and cp06>=" & ChangeTStringToWString(txtCloseDate(0))
      stCon3 = stCon3 & " and ss02>=" & ChangeTStringToWString(txtCloseDate(0))
'edit by nickc 2008/05/14 改成過期都要帶
      'add by nickc 2008/04/24 加入未收款
'      stCon4 = stCon4 & " and BB.rd05>=" & ChangeTStringToWString(txtCloseDate(0))
      'add by nickc 2007/01/23
      stCon1_1 = stCon1_1 & " and np09>=" & ChangeWDateStringToWString(DateAdd("d", 10, ChangeWStringToWDateString(ChangeTStringToWString(txtCloseDate(0)))))
      stCon2_1 = stCon2_1 & " and cp07>=" & ChangeWDateStringToWString(DateAdd("d", 10, ChangeWStringToWDateString(ChangeTStringToWString(txtCloseDate(0)))))
      'Add By Sindy 2011/6/21 未續簽&未回執
      stCon5 = stCon5 & " and cp54>=" & ChangeTStringToWString(txtCloseDate(0))
      '發文日+7天(日曆天)符合查詢本所期限條件
      '2011/7/6 MODIFY BY SONIA 起日抓三個月內
      'stCon6 = stCon6 & " and cp27>=" & ChangeWDateStringToWString(DateAdd("d", -7, ChangeWStringToWDateString(ChangeTStringToWString(txtCloseDate(0)))))
      stCon6 = stCon6 & " and cp27>=" & ChangeWDateStringToWString(DateAdd("m", -3, ChangeWStringToWDateString(ChangeTStringToWString(txtCloseDate(0)))))
   End If
   If txtCloseDate(1) <> "" Then
      If Check5.Value = 0 Then 'Add By Sindy 2014/6/24 +if
         stCon1 = stCon1 & " and np08<=" & ChangeTStringToWString(txtCloseDate(1))
      End If
      stCon9 = stCon9 & " and np23<=" & ChangeTStringToWString(txtCloseDate(1)) 'Added by Morgan 2012/8/21
      
      stCon2 = stCon2 & " and cp06<=" & ChangeTStringToWString(txtCloseDate(1))
      stCon3 = stCon3 & " and ss02<=" & ChangeTStringToWString(txtCloseDate(1))
      'add by nickc 2008/04/24 加入未收款
      'Remove by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
      'stCon4 = stCon4 & " and BB.rd05<=" & ChangeTStringToWString(txtCloseDate(1))
      'stCon8 = stCon8 & " and BB.rd05<=" & ChangeTStringToWString(txtCloseDate(1)) 'Add by Morgan 2011/8/15
      'end 2018/08/22
      'add by nickc 2007/01/23
      stCon1_1 = stCon1_1 & " and np09<=" & ChangeWDateStringToWString(DateAdd("d", 10, ChangeWStringToWDateString(ChangeTStringToWString(txtCloseDate(1)))))
      stCon2_1 = stCon2_1 & " and cp07<=" & ChangeWDateStringToWString(DateAdd("d", 10, ChangeWStringToWDateString(ChangeTStringToWString(txtCloseDate(1)))))
      'Add By Sindy 2011/6/21 未續簽&未回執
      stCon5 = stCon5 & " and cp54<=" & ChangeTStringToWString(txtCloseDate(1))
      '發文日+7天(日曆天)符合查詢本所期限條件
      stCon6 = stCon6 & " and cp27<=" & ChangeWDateStringToWString(DateAdd("d", -7, ChangeWStringToWDateString(ChangeTStringToWString(txtCloseDate(1)))))
   End If
   'Add By Sindy 2014/6/24
   If Check5.Value = 1 Then
      stCon1 = stCon1 & " and ((np08>=" & ChangeTStringToWString(txtCloseDate(0)) & " And np08<=" & ChangeTStringToWString(txtCloseDate(1)) & ") or (np08<" & strSrvDate(1) & " and np09>=" & ChangeTStringToWString(txtCloseDate(0)) & " And np09 <=" & ChangeTStringToWString(txtCloseDate(1)) & "))"
   End If
   '2014/6/24 END
   
   'Add By Sindy 2010/6/17
   If Text1(0) <> "" And Text1(1) <> "" Then
'      stCon4 = stCon4 & " and (cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "') "
'      stConP = stConP & " and (pa01='" & Text1(0) & "' and pa02='" & Text1(1) & "' and pa03='" & Text1(2) & "' and pa04='" & Text1(3) & "') "
'      stConT = stConT & " and (tm01='" & Text1(0) & "' and tm02='" & Text1(1) & "' and tm03='" & Text1(2) & "' and tm04='" & Text1(3) & "') "
'      stConS = stConS & " and (sp01='" & Text1(0) & "' and sp02='" & Text1(1) & "' and sp03='" & Text1(2) & "' and sp04='" & Text1(3) & "') "
'      stConL = stConL & " and (lc01='" & Text1(0) & "' and lc02='" & Text1(1) & "' and lc03='" & Text1(2) & "' and lc04='" & Text1(3) & "') "
'      stConH = stConH & " and (hc01='" & Text1(0) & "' and hc02='" & Text1(1) & "' and hc03='" & Text1(2) & "' and hc04='" & Text1(3) & "') "
      'Modify By Sindy 2010/10/1
      stCon1 = stCon1 & " and (np02='" & Text1(0) & "' and np03='" & Text1(1) & "' and np04='" & Text1(2) & "' and np05='" & Text1(3) & "') "
      stCon9 = stCon9 & " and (np02='" & Text1(0) & "' and np03='" & Text1(1) & "' and np04='" & Text1(2) & "' and np05='" & Text1(3) & "') " 'Added by Morgan 2012/8/21
      
      stCon2 = stCon2 & " and (cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "') "
      stCon4 = stCon4 & " and (cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "') "
      stCon1_1 = stCon1_1 & " and (np02='" & Text1(0) & "' and np03='" & Text1(1) & "' and np04='" & Text1(2) & "' and np05='" & Text1(3) & "') "
      stCon2_1 = stCon2_1 & " and (cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "') "
      'Add By Sindy 2011/6/21 未續簽&未回執
      stCon5 = stCon5 & " and (cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "') "
      stCon6 = stCon6 & " and (cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "') "
      'Add by Morgan 2011/8/15 未收款(新)
      stCon8 = stCon8 & " and (cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "') "
   End If
   
   'add by nickc 2007/01/18
   If Trim(txtCP10) <> "" Then
       stCon1 = stCon1 & " and np07=" & txtCP10 & " "
       stCon9 = stCon9 & " and np07=" & txtCP10 & " " 'Added by Morgan 2012/8/21
       
       stCon2 = stCon2 & " and cp10='" & txtCP10 & "' "
       'add by nickc 2008/04/24 加入未收款
       stCon4 = stCon4 & " and cp10='" & txtCP10 & "' "
       'add by nickc 2007/01/23
       stCon1_1 = stCon1_1 & " and np07=" & txtCP10 & " "
       stCon2_1 = stCon2_1 & " and cp10='" & txtCP10 & "' "
       'Add By Sindy 2011/6/21 未續簽&未回執
       stCon5 = stCon5 & " and cp10='" & txtCP10 & "' "
       stCon6 = stCon6 & " and cp10='" & txtCP10 & "' "
       'Add by Morgan 2011/8/15 未收款(新)
       stCon8 = stCon8 & " and cp10='" & txtCP10 & "' "
   End If
   
   '2009/9/4 add by sonia 未收文部分取消專利處之催審411改為晚上批次通知
   stCon1 = stCon1 & " and np02||np07 not in ('P411','PS411','CFP411','CPS411') "
   'stCon1_1 = stCon1_1 & " and np02||np07 not in ('P411','PS411','CFP411','CPS411') " 'Modify By Sindy 2014/6/24 Mark 因stCon1_1變數使用於trademark
   '2009/9/4 end
   
   stCon9 = stCon9 & " and np02||np07='CFP107' " 'Added by Morgan 2012/8/21
   
   'add by nickc 2007/01/23
   If Trim(txtCU1) <> "" Then
       txtCU1 = Mid(txtCU1 & "000000000", 1, 9)
       txtCU2 = Mid(txtCU2 & "000000000", 1, 9)
       stConP = stConP & " and ((pa26>='" & txtCU1 & "' and pa26<='" & txtCU2 & "') or (pa27>='" & txtCU1 & "' and pa27<='" & txtCU2 & "') or (pa28>='" & txtCU1 & "' and pa28<='" & txtCU2 & "') or (pa29>='" & txtCU1 & "' and pa29<='" & txtCU2 & "') or (pa30>='" & txtCU1 & "' and pa30<='" & txtCU2 & "')) "
       stConT = stConT & " and ((tm23>='" & txtCU1 & "' and tm23<='" & txtCU2 & "') or (tm78>='" & txtCU1 & "' and tm78<='" & txtCU2 & "') or (tm79>='" & txtCU1 & "' and tm79<='" & txtCU2 & "') or (tm80>='" & txtCU1 & "' and tm80<='" & txtCU2 & "') or (tm81>='" & txtCU1 & "' and tm81<='" & txtCU2 & "')) "
       stConS = stConS & " and ((sp08>='" & txtCU1 & "' and sp08<='" & txtCU2 & "') or (sp58>='" & txtCU1 & "' and sp58<='" & txtCU2 & "') or (sp59>='" & txtCU1 & "' and sp59<='" & txtCU2 & "') or (sp65>='" & txtCU1 & "' and sp65<='" & txtCU2 & "') or (sp66>='" & txtCU1 & "' and sp66<='" & txtCU2 & "')) "
       'Modify By Sindy 2011/2/18 增加LC43,LC44,LC45,LC46
       stConL = stConL & " and ((lc11>='" & txtCU1 & "' and lc11<='" & txtCU2 & "') or (lc43>='" & txtCU1 & "' and lc43<='" & txtCU2 & "') or (lc44>='" & txtCU1 & "' and lc44<='" & txtCU2 & "') or (lc45>='" & txtCU1 & "' and lc45<='" & txtCU2 & "') or (lc46>='" & txtCU1 & "' and lc46<='" & txtCU2 & "')) "
       'Modify By Sindy 2011/2/18 增加HC24,HC25,HC26,HC27
       stConH = stConH & " and ((hc05>='" & txtCU1 & "' and hc05<='" & txtCU2 & "') or (hc24>='" & txtCU1 & "' and hc24<='" & txtCU2 & "') or (hc25>='" & txtCU1 & "' and hc25<='" & txtCU2 & "') or (hc26>='" & txtCU1 & "' and hc26<='" & txtCU2 & "') or (hc27>='" & txtCU1 & "' and hc27<='" & txtCU2 & "')) "
       'add by nickc 2008/04/24 加入未收款
       stCon4 = stCon4 & " and ((A0K03>='" & txtCU1 & "' and A0K03<='" & txtCU2 & "')) "
'       stConAcc0k0 = stConAcc0k0 & " and ((A0K03>='" & txtCU1 & "' and A0K03<='" & txtCU2 & "')) "  '2009/11/19 add by sonia
   End If
   
On Error GoTo ErrHnd

'暫存檔規格
'R100123_2:
'   RCP01   VARCHAR2(3)    NULL,
'   RCP02   VARCHAR2(6)    NULL,
'   RCP03   VARCHAR2(1)    NULL,
'   RCP04   VARCHAR2(2)    NULL,
'   RCP09   VARCHAR2(9)    NULL,
'   RCP12   VARCHAR2(3) NULL,          管制部門
'   RCP13   VARCHAR2(6) NULL,
'   RKIND   VARCHAR2(2)    NOT NULL,   分類
'   ID      VARCHAR2(6)    NOT NULL,   strUserNum
'   RCP06   VARCHAR2(12)   NULL,       本所期限
'   REMP    VARCHAR2(6) NULL,          管制人
'   RCP07   VARCHAR2(12)   NULL,       法定期限
'   RNP23   VARCHAR2(12)   NULL,       約定期限
'   RSUBNO  VARCHAR2(50)   NULL,       分所號
'   RCASENAME  VARCHAR2(140)     NULL, 案件名稱
'   RCP10   VARCHAR2(4)    NULL,
'   RCP14   VARCHAR2(6)    NULL,
'   RCP05   VARCHAR2(12)   NULL,
'   RCP27   VARCHAR2(12)   NULL,
'   RAPPID  VARCHAR2(9)    NULL,       申請人1
'   RNATION VARCHAR2(3)    NULL,       申請國家
'   RCASENO VARCHAR2(30)   NULL,       申請案號
'   RNP22   VARCHAR2(10)   NULL,       NP序號
'   RPKEY   VARCHAR2(20)   NULL,       ss01||'-'||ss02||'-'||ss03
   
   'Add By Sindy 2011/6/21 將符合7.未續簽及8.未回執資料存入暫存檔,欲逐筆讀取最新智權人員  '2015/7/16 modify by sonia 加入6.未收款
   strSql = "delete R100123_2 where id='" & strUserNum & "'"
   cnnConnection.Execute strSql, intI
   'Modify by Amy 2014/05/15 +IDXCP54 原:cp01='LA' and cp10='0'
   'Modify By Sindy 2014/6/12 +IDXCP274650
   'Modified by Lydia 2016/12/30 +排除D類收文 and cp27 is null and cp57 is null=> nvl(cp27,0)=0 and nvl(cp57,0)=0 and substr(cp09,1,1)<>'D'
   'modify by sonia 2018/4/26 sqldatet(cp54)改' '||sqldatet(cp54),點本所期限才會正常排序LA-002889
   'Modified by Lydia 2019/04/08 nvl(cp27,0)=0 and nvl(cp57,0)=0 => cp158=0 and cp159=0
   strSql = "insert into R100123_2" & _
           " select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'7','" & strUserNum & "',' '||sqldatet(cp54),cp14 as 管制人,sqldatet(cp07) as 法定期限,'' as 約定期限,hc07 as 分所號,HC06 As 案件名稱,cp10,cp14,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,hc05,'000','' as 申請案號,0 as 序號,'' as PKey from hirecase,caseprogress,staff s2" & _
            " where cp01||cp10='LA0' and cp158=0 and cp159=0 and substr(cp09,1,1)<>'D' and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and cp13=s2.st01(+) " & stCon5 & stConH
   cnnConnection.Execute strSql, intI
   'Modified by Lydia 2016/12/30 若有CP13,則指定Index => IIf(InStr(UCase(stCon6), "CP13") > 0, "/*+ INDEX(CASEPROGRESS IDXCP13051027) */ ", "")
   'modify by sonia 2018/4/26 sqldatet(to_number(to_char(to_date(cp27,'YYYYMMDD')+7,'YYYYMMDD')))改' '||sqldatet(to_number(to_char(to_date(cp27,'YYYYMMDD')+7,'YYYYMMDD')))
   'modify by sonia 2019/7/30 +ACS系統類別
   strSql = "insert into R100123_2" & _
           " select " & IIf(InStr(UCase(stCon6), "CP13") > 0, "/*+ INDEX(CASEPROGRESS IDXCP13051027) */ ", "") & "cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'8','" & strUserNum & "',' '||sqldatet(to_number(to_char(to_date(cp27,'YYYYMMDD')+7,'YYYYMMDD'))),cp14 as 管制人,sqldatet(cp07) as 法定期限,'' as 約定期限,lc16 as 分所號,NVL(NVL(LC05,LC06),LC07) As 案件名稱,cp10,cp14,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,lc11,lc15,'' as 申請案號,0 as 序號,'' as PKey from lawcase,caseprogress,staff s2" & _
            " where cp27>=20110701 and cp01 in('L','CFL','FCL','LIN','ACS','') and cp50 is not null and cp46 is null and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and cp13=s2.st01(+) " & stCon6 & stConL
   cnnConnection.Execute strSql, intI
   'Modified by Lydia 2016/12/30 若有CP13,則指定Index => IIf(InStr(UCase(stCon6), "CP13") > 0, "/*+ INDEX(CASEPROGRESS IDXCP13051027) */ ", "")
   strSql = "insert into R100123_2" & _
           " select " & IIf(InStr(UCase(stCon6), "CP13") > 0, "/*+ INDEX(CASEPROGRESS IDXCP13051027) */ ", "") & "cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'8','" & strUserNum & "',' '||sqldatet(to_number(to_char(to_date(cp27,'YYYYMMDD')+7,'YYYYMMDD'))),cp14 as 管制人,sqldatet(cp07) as 法定期限,'' as 約定期限,hc07 as 分所號,HC06 As 案件名稱,cp10,cp14,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,hc05,'000','' as 申請案號,0 as 序號,'' as PKey from hirecase,caseprogress,staff s2" & _
            " where cp27>=20110701 and cp01='LA' and cp50 is not null and cp46 is null and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and cp13=s2.st01(+) " & stCon6 & stConH
   cnnConnection.Execute strSql, intI
   '2014/6/12 END
   '2015/7/16 add by sonia 加讀離職業務的未續簽資料, 郭章圍應看到楊挺客戶轉給他的資料LA
'   strSql = "insert into R100123_2" & _
'           " select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'7','" & strUserNum & "',sqldatet(cp54),cp14 as 管制人,sqldatet(cp07) as 法定期限,'' as 約定期限,hc07 as 分所號,HC06 As 案件名稱,cp10,cp14,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,hc05,'000','' as 申請案號,0 as 序號,'' as PKey from hirecase,caseprogress,staff s2" & _
'            " where cp01||cp10='LA0' and cp27 is null and cp57 is null and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and cp13=s2.st01(+) and cp13 in ('77010') and cp54>=20150714 and cp54<=20151014 " & stConH
'   cnnConnection.Execute strSql, intI
'   strSql = "insert into R100123_2" & _
'           " select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'8','" & strUserNum & "',sqldatet(to_number(to_char(to_date(cp27,'YYYYMMDD')+7,'YYYYMMDD'))),cp14 as 管制人,sqldatet(cp07) as 法定期限,'' as 約定期限,lc16 as 分所號,NVL(NVL(LC05,LC06),LC07) As 案件名稱,cp10,cp14,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,lc11,lc15,'' as 申請案號,0 as 序號,'' as PKey from lawcase,caseprogress,staff s2" & _
'            " where cp27>=20110701 and cp01 in('L','CFL','FCL','LIN','') and cp50 is not null and cp46 is null and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and cp13=s2.st01(+) and cp13 in ('77010') and cp27>=20150714 and cp27<=20151014 " & stConL
'   cnnConnection.Execute strSql, intI
'   strSql = "insert into R100123_2" & _
'           " select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'8','" & strUserNum & "',sqldatet(to_number(to_char(to_date(cp27,'YYYYMMDD')+7,'YYYYMMDD'))),cp14 as 管制人,sqldatet(cp07) as 法定期限,'' as 約定期限,hc07 as 分所號,HC06 As 案件名稱,cp10,cp14,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,hc05,'000','' as 申請案號,0 as 序號,'' as PKey from hirecase,caseprogress,staff s2" & _
'            " where cp27>=20110701 and cp01='LA' and cp50 is not null and cp46 is null and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and cp13=s2.st01(+) and cp13 in ('77010') and cp27>=20150714 and cp27<=20151014 " & stConH
'   cnnConnection.Execute strSql, intI
   '2015/7/16 end
   
   '2015/7/15 modify by sonia 只需管制人非查
   strSql = "SELECT * FROM R100123_2 WHERE id='" & strUserNum & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
         .MoveFirst
         Do While Not .EOF
            strCP13 = PUB_GetAKindSalesNo(.Fields("rcp01"), .Fields("rcp02"), .Fields("rcp03"), .Fields("rcp04"))
            '2015/7/16 modify by sonia 應改管制人,非智權人員
            'strSql = "Update R100123_2 set rcp12='" & PUB_GetStaffST15(strCP13, "1") & "',rcp13='" & strCP13 & "' where id='" & strUserNum & "' and rcp09='" & .Fields("rcp09") & "'"
            strSql = "Update R100123_2 set rcp12='" & PUB_GetStaffST15(strCP13, "1") & "',REMP='" & strCP13 & "' where id='" & strUserNum & "' and rcp09='" & .Fields("rcp09") & "'"
            cnnConnection.Execute strSql, intI
            .MoveNext
         Loop
      End With
   End If
   '2011/6/21 End
   
   '2015/7/16 add by sonia 因為未收款,未續簽,未回執加讀離職業務資料, 郭章圍應看到楊挺客戶轉給他的資料, 但轉給其他人的則不必看
'   strSql = "delete from R100123_2 where id='" & strUserNum & "' and rkind in ('6','7','8') and (rcp12||''<'S14' or rcp12||''>'S14' or REMP<>'80032')"
'   cnnConnection.Execute strSql, intI
   '2014/6/12 END
   
   Dim tmpCp27SQL As String
   'Modify by Morgan 2009/7/13 +995,996
   'Modified by Morgan 2012/5/23 改寫並加1603
   'MODIFY BY SONIA 2015/5/28 商標加1711通知使用宣誓
   'tmpCp27SQL = "decode(np02||to_char(np07),'L6001',sqldatet(cp27),'FCL6001',sqldatet(cp27),'CFL6001',sqldatet(cp27),'LA6001',sqldatet(cp27),'P999',sqldatet(cp27),'P411',sqldatet(cp27),'P1204',sqldatet(cp27),'P1503',sqldatet(cp27),'PS999',sqldatet(cp27),'PS411',sqldatet(cp27),'PS1204',sqldatet(cp27),'PS1503',sqldatet(cp27),'CFP999',sqldatet(cp27),'CFP411',sqldatet(cp27),'CFP1204',sqldatet(cp27),'CFP1503',sqldatet(cp27),'CPS999',sqldatet(cp27),'CPS411',sqldatet(cp27),'CPS1204',sqldatet(cp27),'CPS1503',sqldatet(cp27),'FCP999',sqldatet(cp27),'FCP411',sqldatet(cp27),'FCP1204',sqldatet(cp27),'FCP1503',sqldatet(cp27),'FG999',sqldatet(cp27),'FG411',sqldatet(cp27),'FG1204',sqldatet(cp27),'FG1503',sqldatet(cp27),'T1403',sqldatet(cp27),'FCT1403',sqldatet(cp27),'FCT312',sqldatet(cp27),'T312',sqldatet(cp27),'CFT312',sqldatet(cp27),'CFT1403',sqldatet(cp27),decode(np07,997,sqldatet(cp27),998,sqldatet(cp27),995,sqldatet(cp27),996,sqldatet(cp27),decode(instr(np02,'T'),0,'',decode(np07,305,sqldatet(cp27),''))))"
   'Modified by Lydia 2016/10/20 TC案+994陸代申請書; 商標案+1701 註冊證
   'modify by sonia 2019/7/30 +ACS系統類別
   tmpCp27SQL = "decode(sign(instr(',L,FCL,CFL,LA,LIN,ACS,',','||np02||',')),1,decode(np07,'6001',sqldatet(cp27),'')" & _
      ",decode(sign(instr(',P,PS,CFP,CPS,FCP,FG,',','||np02||',')),1,decode(sign(instr(',997,998,994,995,996,999,411,1204,1503,1209,1603,',','||np07||',')),1,sqldatet(cp27),'')" & _
      ",decode(sign(instr(',994,997,998,995,996,999,305,1403,1701,1711,312,',','||np07||',')),1,sqldatet(cp27),'')))"
   'End 2012/5/23
   
   Dim tmpKindSql As String
   'Modify by Morgan 2009/7/13 +995,996
   '2009/9/4 modify by sonia 未收文部分取消專利處之催審411改為晚上批次通知
   'MODIFY BY SONIA 2015/5/28 商標加1711通知使用宣誓
   'tmpKindSql = "decode(np02||to_char(np07),'L6001','4','FCL6001','4','CFL6001','4','LA6001','4','P999','4','P411','4','P1204','4','P1503','4','PS999','4','PS411','4','PS1204','4','PS1503','4','CFP999','4','CFP411','4','CFP1204','4','CFP1503','4','CPS999','4','CPS411','4','CPS1204','4','CPS1503','4','FCP999','4','FCP411','4','FCP1204','4','FCP1503','4','FG999','4','FG411','4','FG1204','4','FG1503','4','CFT1403','4','FCT1403','4','T1403','4','CFT312','4','FCT312','4','T312','4','S305','4',decode(np07,997,'4',998,'4',995,'4',996,'4',decode(instr(np02,'T'),0,'2',decode(np07,305,'4','2'))))"
   'Modified by Morgan 2012/5/23 改寫並加1603(專利處411已用條件過濾,應不必再處理,故寫法可與發文日一致)
   'tmpKindSql = "decode(np02||to_char(np07),'L6001','4','FCL6001','4','CFL6001','4','LA6001','4','P999','4','P1204','4','P1503','4','PS999','4','PS1204','4','PS1503','4','CFP999','4','CFP1204','4','CFP1503','4','CPS999','4','CPS1204','4','CPS1503','4','FCP999','4','FCP411','4','FCP1204','4','FCP1503','4','FG999','4','FG411','4','FG1204','4','FG1503','4','CFT1403','4','FCT1403','4','T1403','4','CFT312','4','FCT312','4','T312','4','S305','4',decode(np07,997,'4',998,'4',995,'4',996,'4',decode(instr(np02,'T'),0,'2',decode(np07,305,'4','2'))))"
   'Modified by Lydia 2016/10/20 TC案+994陸代申請書; 商標案+1701 註冊證
   'modify by sonia 2019/7/30 +ACS系統類別
   tmpKindSql = "decode(sign(instr(',L,FCL,CFL,LA,LIN,ACS,',','||np02||',')),1,decode(np07,'6001','4','2')" & _
      ",decode(sign(instr(',P,PS,CFP,CPS,FCP,FG,',','||np02||',')),1,decode(sign(instr(',997,998,994,995,996,999,411,1204,1503,1209,1603,',','||np07||',')),1,'4','2')" & _
      ",decode(sign(instr(',994,997,998,995,996,999,305,1403,1701,1711,312,',','||np07||',')),1,'4','2')))"
   'End 2012/5/23
   '2009/9/4 end
   'Modify By Sindy 2010/01/15 加入PKey
   'Modify By Sindy 2011/3/15 FCT,T,TF延展(102)和第二期(716)專用權須存在(TM17=Y)
   'Modified by Morgan 2012/8/21 +約定期限
   'Add By Sindy 2014/6/12
   If Check1.Value = 1 Or chkNP.Value = 1 Then
      strChkSql = ""
'      If Check1.Value = 1 And chkNP.Value = 1 Then
'         strChkSql = ""
'      Else
'         If chkNP.Value = 1 Then '4.未回覆
'            strChkSql = " and " & tmpKindSql & "='4'"
'         Else '2.未處理
'            strChkSql = " and " & tmpKindSql & "='2'"
'         End If
'      End If
   '2014/6/12 END
      '2.未處理 4.未回覆
'      stCon = "select s2.st06 as 所別,s1.st15 as 管制人部門,np10 as 管制人,cp13 as 收文智權人員," & tmpKindSql & " as 分類,'' as 約定期限,decode(ti01,null,' ','*')||sqldatet(np08) as 本所期限,sqldatet(np09) as 法定期限,np02||'-'||np03||'-'||np04||'-'||np05 as 本所案號,pa47 as 分所號,NVL(NVL(PA05,PA06),PA07) As 案件名稱,NVL(DECODE(PA09,'000',CPM03,CPM04),np07) as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日," & tmpCp27SQL & " as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家,pa11 as 申請案號,np01 as 收文號,np22 as 序號,'' as PKey " & _
'                     " from nextprogress,patent,caseprogress,nation,staff s1,staff s2,customer,casepropertymap,t102inform where np02 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ") " & stCon1 & " and np06 is null and pa57 is null and np01=cp09(+)  and np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+) and np02=cpm01(+) and to_char(np07)=cpm02(+) and pa09=na01(+) and np10=s1.st01(+) and cp13=s2.st01(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) and np01=ti02(+) and np22=ti04(+) " & stConST & stConP & strChkSql
      strSql = "insert into R100123_2" & _
              " select np02,np03,np04,np05,np01,s1.st15 as 管制人部門,cp13," & tmpKindSql & ",'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(np08) as 本所期限,np10 as 管制人,sqldatet(np09) as 法定期限,'' as 約定期限,pa47 as 分所號,NVL(NVL(PA05,PA06),PA07) As 案件名稱,np07 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日," & tmpCp27SQL & " as 發文日,pa26 as 申請人,pa09 as 申請國家,pa11 as 申請案號,np22 as 序號,'' as PKey " & _
                " from nextprogress,patent,caseprogress,staff s1,staff s2,t102inform where np02 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ") and ((" & tmpKindSql & "='2'" & stCon1 & ") or (" & tmpKindSql & "='4'" & Replace(Replace(UCase(stCon1), "NP10", "CP13"), "S1.ST15", "S2.ST15") & ")) and np06 is null and pa57 is null and np01=cp09(+) and np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+) and np10=s1.st01(+) and cp13=s2.st01(+) and np01=ti02(+) and np22=ti04(+) " & stConST & stConP & strChkSql
      cnnConnection.Execute strSql, intI
      'Added by Morgan 2012/8/21
'      stCon = stCon & " union select s2.st06 as 所別,s1.st15 as 管制人部門,np10 as 管制人,cp13 as 收文智權人員,'2' as 分類,sqldatet(np23) as 約定期限,decode(ti01,null,' ','*')||sqldatet(np08) as 本所期限,sqldatet(np09) as 法定期限,np02||'-'||np03||'-'||np04||'-'||np05 as 本所案號,pa47 as 分所號,NVL(NVL(PA05,PA06),PA07) As 案件名稱,NVL(DECODE(PA09,'000',CPM03,CPM04),np07) as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日," & tmpCp27SQL & " as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家,pa11 as 申請案號,np01 as 收文號,np22 as 序號,'' as PKey " & _
'                     " from nextprogress,patent,caseprogress,nation,staff s1,staff s2,customer,casepropertymap,t102inform where np02 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ") " & stCon9 & " and np06 is null and pa57 is null and np01=cp09(+)  and np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+) and np02=cpm01(+) and to_char(np07)=cpm02(+) and pa09=na01(+) and np10=s1.st01(+) and cp13=s2.st01(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) and np01=ti02(+) and np22=ti04(+) " & stConST & stConP & strChkSql
      strSql = "insert into R100123_2" & _
              " select np02,np03,np04,np05,np01,s1.st15 as 管制人部門,cp13,'2','" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(np08) as 本所期限,np10 as 管制人,sqldatet(np09) as 法定期限,sqldatet(np23) as 約定期限,pa47 as 分所號,NVL(NVL(PA05,PA06),PA07) As 案件名稱,np07 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日," & tmpCp27SQL & " as 發文日,pa26 as 申請人,pa09 as 申請國家,pa11 as 申請案號,np22 as 序號,'' as PKey " & _
                " from nextprogress,patent,caseprogress,staff s1,staff s2,t102inform where np02 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ") " & stCon9 & " and np06 is null and pa57 is null and np01=cp09(+) and np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+) and np10=s1.st01(+) and cp13=s2.st01(+) and np01=ti02(+) and np22=ti04(+) " & stConST & stConP & strChkSql
      cnnConnection.Execute strSql, intI
      'end 2012/8/21
'      stCon = stCon & " union select s2.st06 as 所別,s1.st15 as 管制人部門,np10 as 管制人,cp13 as 收文智權人員," & tmpKindSql & " as 分類,'' as 約定期限,decode(ti01,null,' ','*')||sqldatet(np08) as 本所期限,sqldatet(np09) as 法定期限,np02||'-'||np03||'-'||np04||'-'||np05 as 本所案號,tm34 as 分所號,NVL(NVL(tm05,TM06),TM07) As 案件名稱,NVL(DECODE(tm10,'000',CPM03,CPM04),np07) as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日," & tmpCp27SQL & " as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家,tm12 as 申請案號,np01 as 收文號,np22 as 序號,'' as PKey " & _
'                     " from nextprogress,trademark,caseprogress,nation,staff s1,staff s2,customer,casepropertymap,t102inform where np02 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ") " & stCon1 & " and np06 is null and tm29 is null and np01=cp09(+)  and np02=tm01(+) and np03=tm02(+) and np04=tm03(+) and np05=tm04(+) and np02=cpm01(+) and to_char(np07)=cpm02(+) and tm10=na01(+) and np10=s1.st01(+) and cp13=s2.st01(+) and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) and np01=ti02(+)  and np22=ti04(+) " & stConST & stConT & strChkSql & _
'                     " and decode(np02||np07,'T716',tm17,'T102',tm17,'FCT716',tm17,'FCT102',tm17,'TF716',tm17,'TF102',tm17,'Y')='Y' "
      strSql = "insert into R100123_2" & _
              " select np02,np03,np04,np05,np01,s1.st15 as 管制人部門,cp13," & tmpKindSql & ",'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(np08) as 本所期限,np10 as 管制人,sqldatet(np09) as 法定期限,'' as 約定期限,tm34 as 分所號,NVL(NVL(tm05,TM06),TM07) As 案件名稱,np07 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日," & tmpCp27SQL & " as 發文日,tm23 as 申請人,tm10 as 申請國家,tm12 as 申請案號,np22 as 序號,'' as PKey " & _
                " from nextprogress,trademark,caseprogress,staff s1,staff s2,t102inform where np02 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ") and ((" & tmpKindSql & "='2'" & stCon1 & ") or (" & tmpKindSql & "='4'" & Replace(Replace(UCase(stCon1), "NP10", "CP13"), "S1.ST15", "S2.ST15") & ")) and np06 is null and tm29 is null and np01=cp09(+) and np02=tm01(+) and np03=tm02(+) and np04=tm03(+) and np05=tm04(+) and np10=s1.st01(+) and cp13=s2.st01(+) and np01=ti02(+)  and np22=ti04(+) " & stConST & stConT & strChkSql & _
                 " and decode(np02||np07,'T716',tm17,'T102',tm17,'FCT716',tm17,'FCT102',tm17,'TF716',tm17,'TF102',tm17,'Y')='Y' "
      cnnConnection.Execute strSql, intI
      '2007/12/12 modify by sonia 加入FCT208
'      stCon = stCon & " union select s2.st06 as 所別,s1.st15 as 管制人部門,np10 as 管制人,cp13 as 收文智權人員," & tmpKindSql & " as 分類,'' as 約定期限,decode(ti01,null,' ','*')||sqldatet(np08) as 本所期限,sqldatet(np09) as 法定期限,np02||'-'||np03||'-'||np04||'-'||np05 as 本所案號,tm34 as 分所號,NVL(NVL(tm05,TM06),TM07) As 案件名稱,NVL(DECODE(tm10,'000',CPM03,CPM04),np07) as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日," & tmpCp27SQL & " as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家,tm12 as 申請案號,np01 as 收文號,np22 as 序號,'' as PKey " & _
'                     " from nextprogress,trademark,caseprogress,nation,staff s1,staff s2,customer,casepropertymap,t102inform where np02 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ") " & stCon1_1 & " and np06 is null and tm29 is null and np01=cp09(+)  and np02=tm01(+) and np03=tm02(+) and np04=tm03(+) and np05=tm04(+) and np02=cpm01(+) and to_char(np07)=cpm02(+) and tm10=na01(+) and np10=s1.st01(+) and cp13=s2.st01(+) and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) and np01=ti02(+)  and np22=ti04(+) and np02||to_char(np07) in ('CFT102','CFT105','FCT208')  " & stConST & stConT & strChkSql & _
'                     " and decode(np02||np07,'T716',tm17,'T102',tm17,'FCT716',tm17,'FCT102',tm17,'TF716',tm17,'TF102',tm17,'Y')='Y' "
      strSql = "insert into R100123_2" & _
              " select np02,np03,np04,np05,np01,s1.st15 as 管制人部門,cp13," & tmpKindSql & ",'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(np08) as 本所期限,np10 as 管制人,sqldatet(np09) as 法定期限,'' as 約定期限,tm34 as 分所號,NVL(NVL(tm05,TM06),TM07) As 案件名稱,np07 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日," & tmpCp27SQL & " as 發文日,tm23 as 申請人,tm10 as 申請國家,tm12 as 申請案號,np22 as 序號,'' as PKey " & _
                " from nextprogress,trademark,caseprogress,staff s1,staff s2,t102inform where np02 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ") " & stCon1_1 & " and np06 is null and tm29 is null and np01=cp09(+) and np02=tm01(+) and np03=tm02(+) and np04=tm03(+) and np05=tm04(+) and np10=s1.st01(+) and cp13=s2.st01(+) and np01=ti02(+)  and np22=ti04(+) and np02||to_char(np07) in ('CFT102','CFT105','FCT208')  " & stConST & stConT & strChkSql & _
                 " and decode(np02||np07,'T716',tm17,'T102',tm17,'FCT716',tm17,'FCT102',tm17,'TF716',tm17,'TF102',tm17,'Y')='Y' "
      cnnConnection.Execute strSql, intI
'      stCon = stCon & " union select s2.st06 as 所別,s1.st15 as 管制人部門,np10 as 管制人,cp13 as 收文智權人員," & tmpKindSql & " as 分類,'' as 約定期限,decode(ti01,null,' ','*')||sqldatet(np08) as 本所期限,sqldatet(np09) as 法定期限,np02||'-'||np03||'-'||np04||'-'||np05 as 本所案號,sp28 as 分所號,NVL(NVL(SP05,SP06),SP07) As 案件名稱,NVL(DECODE(sp09,'000',CPM03,CPM04),np07) as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日," & tmpCp27SQL & " as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家,sp11 as 申請案號,np01 as 收文號,np22 as 序號,'' as PKey " & _
'                     " from nextprogress,servicepractice,caseprogress,nation,staff s1,staff s2,customer,casepropertymap,t102inform where np02 in (" & SQLGrpStr(GetAllSysKind(systemkind), 5) & ") " & stCon1 & " and np06 is null and sp15 is null and np01=cp09(+)  and np02=sp01(+) and np03=sp02(+) and np04=sp03(+) and np05=sp04(+) and np02=cpm01(+) and to_char(np07)=cpm02(+) and sp09=na01(+) and np10=s1.st01(+) and cp13=s2.st01(+) and substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+) and np01=ti02(+)  and np22=ti04(+) " & stConST & stConS & strChkSql
      strSql = "insert into R100123_2" & _
              " select np02,np03,np04,np05,np01,s1.st15 as 管制人部門,cp13," & tmpKindSql & ",'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(np08) as 本所期限,np10 as 管制人,sqldatet(np09) as 法定期限,'' as 約定期限,sp28 as 分所號,NVL(NVL(SP05,SP06),SP07) As 案件名稱,np07 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日," & tmpCp27SQL & " as 發文日,sp08 as 申請人,sp09 as 申請國家,sp11 as 申請案號,np22 as 序號,'' as PKey " & _
                " from nextprogress,servicepractice,caseprogress,staff s1,staff s2,t102inform where np02 in (" & SQLGrpStr(GetAllSysKind(systemkind), 5) & ") and ((" & tmpKindSql & "='2'" & stCon1 & ") or (" & tmpKindSql & "='4'" & Replace(Replace(UCase(stCon1), "NP10", "CP13"), "S1.ST15", "S2.ST15") & ")) and np06 is null and sp15 is null and np01=cp09(+) and np02=sp01(+) and np03=sp02(+) and np04=sp03(+) and np05=sp04(+) and np10=s1.st01(+) and cp13=s2.st01(+) and np01=ti02(+)  and np22=ti04(+) " & stConST & stConS & strChkSql
      cnnConnection.Execute strSql, intI
'      stCon = stCon & " union select s2.st06 as 所別,s1.st15 as 管制人部門,np10 as 管制人,cp13 as 收文智權人員," & tmpKindSql & " as 分類,'' as 約定期限,decode(ti01,null,' ','*')||sqldatet(np08) as 本所期限,sqldatet(np09) as 法定期限,np02||'-'||np03||'-'||np04||'-'||np05 as 本所案號,lc16 as 分所號,NVL(NVL(LC05,LC06),LC07) As 案件名稱,NVL(DECODE(lc15,'000',CPM03,CPM04),np07) as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日," & tmpCp27SQL & " as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家,'' as 申請案號,np01 as 收文號,np22 as 序號,'' as PKey " & _
'                     " from nextprogress,lawcase,caseprogress,nation,staff s1,staff s2,customer,casepropertymap,t102inform where np02 in (" & SQLGrpStr(GetAllSysKind(systemkind), 3) & ") " & stCon1 & " and np06 is null and lc08 is null and np01=cp09(+)  and np02=lc01(+) and np03=lc02(+) and np04=lc03(+) and np05=lc04(+) and np02=cpm01(+) and to_char(np07)=cpm02(+) and lc15=na01(+) and np10=s1.st01(+) and cp13=s2.st01(+) and substr(lc11,1,8)=cu01(+) and substr(lc11,9,1)=cu02(+) and np01=ti02(+)  and np22=ti04(+) " & stConST & stConL & strChkSql
      strSql = "insert into R100123_2" & _
              " select np02,np03,np04,np05,np01,s1.st15 as 管制人部門,cp13," & tmpKindSql & ",'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(np08) as 本所期限,np10 as 管制人,sqldatet(np09) as 法定期限,'' as 約定期限,lc16 as 分所號,NVL(NVL(LC05,LC06),LC07) As 案件名稱,np07 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日," & tmpCp27SQL & " as 發文日,lc11 as 申請人,lc15 as 申請國家,'' as 申請案號,np22 as 序號,'' as PKey " & _
                " from nextprogress,lawcase,caseprogress,staff s1,staff s2,t102inform where np02 in (" & SQLGrpStr(GetAllSysKind(systemkind), 3) & ") and ((" & tmpKindSql & "='2'" & stCon1 & ") or (" & tmpKindSql & "='4'" & Replace(Replace(UCase(stCon1), "NP10", "CP13"), "S1.ST15", "S2.ST15") & ")) and np06 is null and lc08 is null and np01=cp09(+) and np02=lc01(+) and np03=lc02(+) and np04=lc03(+) and np05=lc04(+) and np10=s1.st01(+) and cp13=s2.st01(+) and np01=ti02(+)  and np22=ti04(+) " & stConST & stConL & strChkSql
      cnnConnection.Execute strSql, intI
'      stCon = stCon & " union select s2.st06 as 所別,s1.st15 as 管制人部門,np10 as 管制人,cp13 as 收文智權人員," & tmpKindSql & " as 分類,'' as 約定期限,decode(ti01,null,' ','*')||sqldatet(np08) as 本所期限,sqldatet(np09) as 法定期限,np02||'-'||np03||'-'||np04||'-'||np05 as 本所案號,hc07 as 分所號,HC06                                        As 案件名稱,NVL(CPM03,np07)                                             as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日," & tmpCp27SQL & " as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家,'' as 申請案號,np01 as 收文號,np22 as 序號,'' as PKey " & _
'                     " from nextprogress,hirecase,caseprogress,nation,staff s1,staff s2,customer,casepropertymap,t102inform where np02 in (" & SQLGrpStr(GetAllSysKind(systemkind), 4) & ") " & stCon1 & " and np06 is null and hc09 is null and np01=cp09(+)  and np02=hc01(+) and np03=hc02(+) and np04=hc03(+) and np05=hc04(+) and np02=cpm01(+) and to_char(np07)=cpm02(+) and '000'=na01(+) and np10=s1.st01(+) and cp13=s2.st01(+) and substr(hc05,1,8)=cu01(+) and substr(hc05,9,1)=cu02(+) and np01=ti02(+)  and np22=ti04(+) " & stConST & stConH & strChkSql
      strSql = "insert into R100123_2" & _
              " select np02,np03,np04,np05,np01,s1.st15 as 管制人部門,cp13," & tmpKindSql & ",'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(np08) as 本所期限,np10 as 管制人,sqldatet(np09) as 法定期限,'' as 約定期限,hc07 as 分所號,HC06 As 案件名稱,np07 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日," & tmpCp27SQL & " as 發文日,hc05 as 申請人,'000' as 申請國家,'' as 申請案號,np22 as 序號,'' as PKey " & _
                " from nextprogress,hirecase,caseprogress,staff s1,staff s2,t102inform where np02 in (" & SQLGrpStr(GetAllSysKind(systemkind), 4) & ") and ((" & tmpKindSql & "='2'" & stCon1 & ") or (" & tmpKindSql & "='4'" & Replace(Replace(UCase(stCon1), "NP10", "CP13"), "S1.ST15", "S2.ST15") & ")) and np06 is null and hc09 is null and np01=cp09(+) and np02=hc01(+) and np03=hc02(+) and np04=hc03(+) and np05=hc04(+) and np10=s1.st01(+) and cp13=s2.st01(+) and np01=ti02(+)  and np22=ti04(+) " & stConST & stConH & strChkSql
      cnnConnection.Execute strSql, intI
      
      If chkNP.Value = 0 Then '無4.未回覆
         strSql = "delete from R100123_2 where id='" & strUserNum & "' and rkind='4'"
         cnnConnection.Execute strSql, intI
      ElseIf chkNP.Value = 1 And Check5.Value = 1 Then '有4.未回覆 刪除含此期間法定期限案件(逾本所期限)
         strSql = "delete from R100123_2 where id='" & strUserNum & "' and rkind='4'" & _
                    " and to_char(to_date(replace(replace(rcp06,'/',''),'*','')+19110000,'YYYYMMDD'),'YYYYMMDD')<" & strSrvDate(1) & " and to_char(to_date(replace(replace(rcp07,'/',''),'*','')+19110000,'YYYYMMDD'),'YYYYMMDD')>=" & DBDATE(txtCloseDate(0)) & " And to_char(to_date(replace(replace(rcp07,'/',''),'*','')+19110000,'YYYYMMDD'),'YYYYMMDD')<=" & DBDATE(txtCloseDate(1))
         cnnConnection.Execute strSql, intI
      End If
      'Add By Sindy 2019/9/26
      If Check1.Value = 0 Then '無2.未處理
         strSql = "delete from R100123_2 where id='" & strUserNum & "' and rkind='2'"
         cnnConnection.Execute strSql, intI
      End If
      '2019/9/26 END
   End If
   
   '9.未函知
   If Check1.Value = 1 Then
      'Add By Sindy 2015/3/2 9.未函知
      'Modify By Sindy 2015/3/27 + and cp30=RNP22 因為像延展下一程式就有相關文號的資料存在,如.第二期註冊費 T-174377
      'Modified by Morgan 2016/4/18 +119
      'modify by sonia 2017/11/24 +930
      strSql = "update R100123_2 set rkind='9' where id='" & strUserNum & "' and rkind='2' and rcp01 in('CFP','P') and rCP10 in ('605','606','607','416','119','930')" & _
                 " and not exists (select cp09 from caseprogress where cp43=rcp09 and cp10='1913' and cp30=RNP22)"
      cnnConnection.Execute strSql, intI
      'Modify By Sindy 2015/3/25 將102延展另外判斷,因1725通知期限是104/3月才加的程式,所以通知延展之前就有發出去了,因此加本所期限控管
'      strSql = "update R100123_2 set rkind='9' where id='" & strUserNum & "' and rkind='2' and rcp01 in('T','TB','TF') and rCP10 in ('102','105','109','702','708','716')" & _
'                 " and not exists (select cp09 from caseprogress where cp43=rcp09 and cp10='1725')"
      strSql = "update R100123_2 set rkind='9' where id='" & strUserNum & "' and rkind='2' and rcp01 in('T','TB','TF') and rCP10 in ('105','109','702','708','716')" & _
                 " and not exists (select cp09 from caseprogress where cp43=rcp09 and cp10='1725' and cp30=RNP22)"
      cnnConnection.Execute strSql, intI
      strSql = "update R100123_2 set rkind='9' where id='" & strUserNum & "' and rkind='2' and rcp01 in('T','TF') and rCP10 in ('102')" & _
                 " and not exists (select cp09 from caseprogress where cp43=rcp09 and cp10='1725' and cp30=RNP22)" & _
                 " and to_number(replace(replace(RCP06,'/',''),'*',''))>=1050401"
      cnnConnection.Execute strSql, intI
      '2015/3/25 END
      strSql = "update R100123_2 set rkind='9' where id='" & strUserNum & "' and rkind='2' and rcp01 in('CFT') and rCP10 in ('102')" & _
                 " and not exists (select cp09 from caseprogress where cp43=rcp09 and cp10='1717' and cp30=RNP22)"
      cnnConnection.Execute strSql, intI
      '1723.本所通知使用宣誓
      '1711.通知使用宣誓
      'Modify By Sindy 2017/2/20 調整1723,1711的串法
      strSql = "update R100123_2 set rkind='9' where id='" & strUserNum & "' and rkind='2' and rcp01 in('CFT') and rCP10 in ('105')" & _
                 " and not exists (select cp09 from caseprogress where cp43=rcp09 and cp10='1723' and cp30=RNP22)" & _
                 " and not exists (select cp09 from caseprogress where cp09=rcp09 and cp10='1711')"
      cnnConnection.Execute strSql, intI
      '2015/3/2 END
   End If
   
   'Add By Sindy 2014/6/12
   If Check3.Value = 1 Or Check4.Value = 1 Then
      If Check3.Value = 1 And Check4.Value = 1 Then
         strChkSql = ""
      Else
         'modify by sonia 2015/10/29 法務案件之通知開庭9001改以'未發文'顯示
         'If Check3.Value = 1 Then '5.未通知
         '   strChkSql = " and substr(cp09,1,1)='C'"
         'Else '1.未發文
         '   strChkSql = " and substr(cp09,1,1)<>'C'"
         'End If
         If Check3.Value = 1 Then '5.未通知
            strChkSql = " and substr(cp09,1,1)='C' and cp10<>'9001'"
         Else '1.未發文
            strChkSql = " and (substr(cp09,1,1)<>'C' or cp10='9001')"
         End If
         'end 2015/10/29
      End If
   '2014/6/12 END
      '1.未發文 5.未通知
'      stCon = stCon & " union select s2.st06 as 所別,cp12 as 管制人部門,cp14 as 管制人,cp13 as 收文智權人員,decode(substr(cp09,1,1),'C','5','1') as 分類,'' as 約定期限,decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,pa47 as 分所號,NVL(NVL(PA05,PA06),PA07) As 案件名稱,NVL(DECODE(PA09,'000',CPM03,CPM04),cp10) as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家,pa11 as 申請案號,cp09 as 收文號,0 as 序號,'' as PKey " & _
'                     " from patent,caseprogress,nation,staff s2,customer,casepropertymap,t102inform where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ") " & stCon2 & strChkSql & " and cp27 is null and cp57 is null and pa57 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp01=cpm01(+) and cp10=cpm02(+) and pa09=na01(+) and cp13= s2.st01(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) and cp09=ti02(+) " & stConST & stConP
      'Modified by Lydia 2016/12/30 +排除D類收文 cp27 is null and cp57 is null => nvl(cp27,0)=0 and nvl(cp57,0)=0 and substr(cp09,1,1)<>'D'
      'Modified by Lydia 2019/04/08 nvl(cp27,0)=0 and nvl(cp57,0)=0 => cp158=0 and cp159=0
      strSql = "insert into R100123_2" & _
              " select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,decode(substr(cp09,1,1),'C','5','1') as 分類,'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,cp14 as 管制人,sqldatet(cp07) as 法定期限,'' as 約定期限,pa47 as 分所號,NVL(NVL(PA05,PA06),PA07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,pa26 as 申請人,pa09 as 申請國家,pa11 as 申請案號,0 as 序號,'' as PKey " & _
                " from patent,caseprogress,staff s2,t102inform where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ") " & stCon2 & strChkSql & " and cp158=0 and cp159=0 and substr(cp09,1,1)<>'D' and pa57 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp13= s2.st01(+) and cp09=ti02(+) " & stConST & stConP
      cnnConnection.Execute strSql, intI
'      stCon = stCon & " union select s2.st06 as 所別,cp12 as 管制人部門,cp14 as 管制人,cp13 as 收文智權人員,decode(substr(cp09,1,1),'C','5','1') as 分類,'' as 約定期限,decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,tm34 as 分所號,NVL(NVL(tm05,TM06),TM07) As 案件名稱,NVL(DECODE(tm10,'000',CPM03,CPM04),cp10) as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家,tm12 as 申請案號,cp09 as 收文號,0 as 序號,'' as PKey " & _
'                     " from trademark,caseprogress,nation,staff s2,customer,casepropertymap,t102inform where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ") " & stCon2 & strChkSql & " and cp27 is null and cp57 is null and tm29 is null and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp01=cpm01(+) and cp10=cpm02(+) and tm10=na01(+) and cp13=s2.st01(+) and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) and cp09=ti02(+) " & stConST & stConT
      'Modified by Lydia 2016/12/30 +排除D類收文 cp27 is null and cp57 is null => nvl(cp27,0)=0 and nvl(cp57,0)=0 and substr(cp09,1,1)<>'D'
      'Modified by Lydia 2019/04/08 nvl(cp27,0)=0 and nvl(cp57,0)=0 => cp158=0 and cp159=0
      strSql = "insert into R100123_2" & _
              " select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,decode(substr(cp09,1,1),'C','5','1') as 分類,'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,cp14 as 管制人,sqldatet(cp07) as 法定期限,'' as 約定期限,tm34 as 分所號,NVL(NVL(tm05,TM06),TM07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,tm23 as 申請人,tm10 as 申請國家,tm12 as 申請案號,0 as 序號,'' as PKey " & _
                " from trademark,caseprogress,staff s2,t102inform where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ") " & stCon2 & strChkSql & " and cp158=0 and cp159=0 and substr(cp09,1,1)<>'D' and tm29 is null and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp13=s2.st01(+) and cp09=ti02(+) " & stConST & stConT
      cnnConnection.Execute strSql, intI
'      stCon = stCon & " union select s2.st06 as 所別,cp12 as 管制人部門,cp14 as 管制人,cp13 as 收文智權人員,decode(substr(cp09,1,1),'C','5','1') as 分類,'' as 約定期限,decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,tm34 as 分所號,NVL(NVL(tm05,TM06),TM07) As 案件名稱,NVL(DECODE(tm10,'000',CPM03,CPM04),cp10) as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家,tm12 as 申請案號,cp09 as 收文號,0 as 序號,'' as PKey " & _
'                     " from trademark,caseprogress,nation,staff s2,customer,casepropertymap,t102inform where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ") " & stCon2_1 & strChkSql & " and cp27 is null and cp57 is null and tm29 is null and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp01=cpm01(+) and cp10=cpm02(+) and tm10=na01(+) and cp13=s2.st01(+) and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) and cp09=ti02(+) and cp01||cp10 in ('CFT102','CFT105','FCT208') " & stConST & stConT
      'Modified by Lydia 2016/12/30 +排除D類收文 cp27 is null and cp57 is null => nvl(cp27,0)=0 and nvl(cp57,0)=0 and substr(cp09,1,1)<>'D'
      'Modified by Lydia 2019/04/08 nvl(cp27,0)=0 and nvl(cp57,0)=0 => cp158=0 and cp159=0
      strSql = "insert into R100123_2" & _
              " select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,decode(substr(cp09,1,1),'C','5','1') as 分類,'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,cp14 as 管制人,sqldatet(cp07) as 法定期限,'' as 約定期限,tm34 as 分所號,NVL(NVL(tm05,TM06),TM07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,tm23 as 申請人,tm10 as 申請國家,tm12 as 申請案號,0 as 序號,'' as PKey " & _
                " from trademark,caseprogress,staff s2,t102inform where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ") " & stCon2_1 & strChkSql & " and cp158=0 and cp159=0 and substr(cp09,1,1)<>'D' and tm29 is null and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp13=s2.st01(+) and cp09=ti02(+) and cp01||cp10 in ('CFT102','CFT105','FCT208') " & stConST & stConT
      cnnConnection.Execute strSql, intI
'      stCon = stCon & " union select s2.st06 as 所別,cp12 as 管制人部門,cp14 as 管制人,cp13 as 收文智權人員,decode(substr(cp09,1,1),'C','5','1') as 分類,'' as 約定期限,decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,sp28 as 分所號,NVL(NVL(SP05,SP06),SP07) As 案件名稱,NVL(DECODE(sp09,'000',CPM03,CPM04),cp10) as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家,sp11 as 申請案號,cp09 as 收文號,0 as 序號,'' as PKey " & _
'                     " from servicepractice,caseprogress,nation,staff s2,customer,casepropertymap,t102inform where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 5) & ") " & stCon2 & strChkSql & " and cp27 is null and cp57 is null and sp15 is null and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and cp01=cpm01(+) and cp10=cpm02(+) and sp09=na01(+) and cp13=s2.st01(+) and substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+) and cp09=ti02(+) " & stConST & stConS
      'Modified by Lydia 2016/12/30 +排除D類收文 cp27 is null and cp57 is null => nvl(cp27,0)=0 and nvl(cp57,0)=0 and substr(cp09,1,1)<>'D'
      'Modified by Lydia 2019/04/08 nvl(cp27,0)=0 and nvl(cp57,0)=0 => cp158=0 and cp159=0
      strSql = "insert into R100123_2" & _
              " select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,decode(substr(cp09,1,1),'C','5','1') as 分類,'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,cp14 as 管制人,sqldatet(cp07) as 法定期限,'' as 約定期限,sp28 as 分所號,NVL(NVL(SP05,SP06),SP07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,sp08 as 申請人,sp09 as 申請國家,sp11 as 申請案號,0 as 序號,'' as PKey " & _
                " from servicepractice,caseprogress,staff s2,t102inform where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 5) & ") " & stCon2 & strChkSql & " and cp158=0 and cp159=0 and substr(cp09,1,1)<>'D' and sp15 is null and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and cp13=s2.st01(+) and cp09=ti02(+) " & stConST & stConS
      cnnConnection.Execute strSql, intI
'      stCon = stCon & " union select s2.st06 as 所別,cp12 as 管制人部門,cp14 as 管制人,cp13 as 收文智權人員,decode(substr(cp09,1,1),'C','5','1') as 分類,'' as 約定期限,decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,lc16 as 分所號,NVL(NVL(LC05,LC06),LC07) As 案件名稱,NVL(DECODE(lc15,'000',CPM03,CPM04),cp10) as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家,'' as 申請案號,cp09 as 收文號,0 as 序號,'' as PKey " & _
'                     " from lawcase,caseprogress,nation,staff s2,customer,casepropertymap,t102inform where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 3) & ") " & stCon2 & strChkSql & " and cp27 is null and cp57 is null and lc08 is null and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and cp01=cpm01(+) and cp10=cpm02(+) and lc15=na01(+) and cp13=s2.st01(+) and substr(lc11,1,8)=cu01(+) and substr(lc11,9,1)=cu02(+) and cp09=ti02(+) " & stConST & stConL
      'modify by sonia 2015/10/29 法務案件之通知開庭9001改以'未發文'顯示
      'strSql = "insert into R100123_2" & _
      '        " select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,decode(substr(cp09,1,1),'C','5','1') as 分類,'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,cp14 as 管制人,sqldatet(cp07) as 法定期限,'' as 約定期限,lc16 as 分所號,NVL(NVL(LC05,LC06),LC07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,lc11 as 申請人,lc15 as 申請國家,'' as 申請案號,0 as 序號,'' as PKey " & _
      '          " from lawcase,caseprogress,staff s2,t102inform where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 3) & ") " & stCon2 & strChkSql & " and cp27 is null and cp57 is null and lc08 is null and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and cp13=s2.st01(+) and cp09=ti02(+) " & stConST & stConL
      'Modified by Lydia 2016/12/30 +排除D類收文 cp27 is null and cp57 is null => nvl(cp27,0)=0 and nvl(cp57,0)=0 and substr(cp09,1,1)<>'D'
      'Modified by Lydia 2019/04/08 nvl(cp27,0)=0 and nvl(cp57,0)=0 => cp158=0 and cp159=0
      strSql = "insert into R100123_2" & _
              " select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,decode(cp10,'9001','1',decode(substr(cp09,1,1),'C','5','1')) as 分類,'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,cp14 as 管制人,sqldatet(cp07) as 法定期限,'' as 約定期限,lc16 as 分所號,NVL(NVL(LC05,LC06),LC07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,lc11 as 申請人,lc15 as 申請國家,'' as 申請案號,0 as 序號,'' as PKey " & _
                " from lawcase,caseprogress,staff s2,t102inform where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 3) & ") " & stCon2 & strChkSql & " and cp158=0 and cp159=0 and substr(cp09,1,1)<>'D' and lc08 is null and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and cp13=s2.st01(+) and cp09=ti02(+) " & stConST & stConL
      'end 2015/10/29
      cnnConnection.Execute strSql, intI
'      stCon = stCon & " union select s2.st06 as 所別,cp12 as 管制人部門,cp14 as 管制人,cp13 as 收文智權人員,decode(substr(cp09,1,1),'C','5','1') as 分類,'' as 約定期限,decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,hc07 as 分所號,HC06                                        As 案件名稱,NVL(CPM03,cp10)                                             as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家,'' as 申請案號,cp09 as 收文號,0 as 序號,'' as PKey " & _
'                     " from hirecase,caseprogress,nation,staff s2,customer,casepropertymap,t102inform where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 4) & ") " & stCon2 & strChkSql & " and cp27 is null and cp57 is null and hc09 is null and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and cp01=cpm01(+) and cp10=cpm02(+) and '000'=na01(+) and cp13=s2.st01(+) and substr(hc05,1,8)=cu01(+) and substr(hc05,9,1)=cu02(+) and cp09=ti02(+) " & stConST & stConH
      'Modified by Lydia 2016/12/30 +排除D類收文 cp27 is null and cp57 is null => nvl(cp27,0)=0 and nvl(cp57,0)=0 and substr(cp09,1,1)<>'D'
      'Modified by Lydia 2019/04/08 nvl(cp27,0)=0 and nvl(cp57,0)=0 => cp158=0 and cp159=0
      strSql = "insert into R100123_2" & _
              " select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,decode(substr(cp09,1,1),'C','5','1') as 分類,'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,cp14 as 管制人,sqldatet(cp07) as 法定期限,'' as 約定期限,hc07 as 分所號,HC06 As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,hc05 as 申請人,'000' as 申請國家,'' as 申請案號,0 as 序號,'' as PKey " & _
                " from hirecase,caseprogress,staff s2,t102inform where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 4) & ") " & stCon2 & strChkSql & " and cp158=0 and cp159=0 and substr(cp09,1,1)<>'D' and hc09 is null and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and cp13=s2.st01(+) and cp09=ti02(+) " & stConST & stConH
      cnnConnection.Execute strSql, intI
   End If
   
'   'Add By Sindy 2011/6/21 讀取7.未續簽及8.未回執資料
'   stCon = stCon & " union select s2.st06 as 所別,rcp12 as 管制人部門,cp14 as 管制人,rcp13 as 收文智權人員,rkind as 分類,'' as 約定期限,' '||sqldatet(rcp06) as 本所期限,sqldatet(cp07) as 法定期限,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,hc07 as 分所號,HC06                     As 案件名稱,NVL(CPM03,cp10)                          as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家,'' as 申請案號,cp09 as 收文號,0 as 序號,'' as PKey " & _
'                  " from R100123_2,hirecase,caseprogress,nation,staff s2,customer,casepropertymap" & _
'                  " where id='" & strUserNum & "' and rcp01=hc01(+) and rcp02=hc02(+) and rcp03=hc03(+) and rcp04=hc04(+) and rcp09=cp09(+) and cp01=cpm01(+) and cp10=cpm02(+) and '000'=na01(+) and rcp13=s2.st01(+) and substr(hc05,1,8)=cu01(+) and substr(hc05,9,1)=cu02(+) " & stCon7
'   stCon = stCon & " union select s2.st06 as 所別,rcp12 as 管制人部門,cp14 as 管制人,rcp13 as 收文智權人員,rkind as 分類,'' as 約定期限,' '||sqldatet(rcp06) as 本所期限,sqldatet(cp07) as 法定期限,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,lc16 as 分所號,NVL(NVL(LC05,LC06),LC07) As 案件名稱,NVL(DECODE(lc15,'000',CPM03,CPM04),cp10) as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家,'' as 申請案號,cp09 as 收文號,0 as 序號,'' as PKey " & _
'                  " from R100123_2,lawcase,caseprogress,nation,staff s2,customer,casepropertymap" & _
'                  " where id='" & strUserNum & "' and rcp01=lc01(+) and rcp02=lc02(+) and rcp03=lc03(+) and rcp04=lc04(+) and rcp09=cp09(+) and cp01=cpm01(+) and cp10=cpm02(+) and lc15=na01(+) and rcp13=s2.st01(+) and substr(lc11,1,8)=cu01(+) and substr(lc11,9,1)=cu02(+) " & stCon7
   
   'Add By Sindy 2010/10/1 若有輸入查詢條件為系統類別,本所案號,案件性質者,不查詢行事曆資料
   If systemkind <> "ALL" Or (Text1(0) <> "" And Text1(1) <> "") Or txtCP10 <> "" Then
      '不查詢行事曆資料
   '2010/10/1 End
   Else
'      stCon = stCon & " union select s2.st06 as 所別,s2.st15 as 管制人部門,ss01 as 管制人,ss01 as 收文智權人員,'3' as 分類,'' as 約定期限,' '||sqldatet(ss02) as 本所期限,'' as 法定期限,'' as 本所案號,'' as 分所號,ss04 As 案件名稱,'' as 案件性質,'' as 承辦人,'' as 收文日,'' as 發文日,'' as 申請人,'' as 申請國家,'' as 申請案號,'' as 收文號,0 as 序號,ss01||'-'||ss02||'-'||ss03 as PKey " & _
'                     " from staff_schedule,staff s2 where ss01=s2.st01(+) " & stConST & stCon3
      'Modified by Morgan 2015/8/31 案件名稱要限制長度
      strSql = "insert into R100123_2" & _
              " select '' as cp01,'' as cp02,'' as cp03,'' as cp04,'' as cp09,s2.st15 as 管制人部門,ss01,'3' as 分類,'" & strUserNum & "',' '||sqldatet(ss02) as 本所期限,ss01 as 管制人,'' as 法定期限,'' as 約定期限,'' as 分所號,substrb(ss04,1,140) As 案件名稱,'' as 案件性質,'' as 承辦人,'' as 收文日,'' as 發文日,'' as 申請人,'' as 申請國家,'' as 申請案號,0 as 序號,ss01||'-'||ss02||'-'||ss03 as PKey " & _
                " from staff_schedule,staff s2 where ss01=s2.st01(+) " & stConST & stCon3
      cnnConnection.Execute strSql, intI
   End If
   
'Modify by Morgan 2011/8/15 改語法,ReceivablesDay 加 已收款(RD06) 欄位
'
'   'add by nickc 2008/04/24 加入未收款
'   Dim oDiffTB As String
''Modify by Morgan 2010/1/29 改寫暫存表R100123
''   oDiffTB = "  select a0k01 oKey,SUM(nvl(a0k06,0)+nvl(a0k07,0)-PAY) diff FROM(  "
''   oDiffTB = oDiffTB & "  SELECT A0K01,nvl(A0K06,0) as a0k06,nvl(A0K07,0) as a0k07,sum(nvl(a1u04,0))+sum(nvl(a1u07,0))-sum(nvl(a1u08,0))+sum(nvl(a1u05,0))+sum(nvl(a1u09,0))-sum(nvl(a1u10,0)) PAY "
''   '2009/11/19 modify by sonia 調整未收款語法,加入業務區智權人員條件
''   'oDiffTB = oDiffTB & "  from acc0k0,ACC1U0 WHERE a0k02 <" & strSrvDate(2) & " AND (a0k09 is null or a0k09 = 0) AND A0K01=A1U02(+) "
''   oDiffTB = oDiffTB & " from acc0k0,ACC1U0 WHERE a0k02 <" & strSrvDate(2) & " AND (a0k09 is null or a0k09 = 0) " & stConAcc0k0 & " AND A0K01=A1U02(+) "
''   '2009/11/19 end
''   oDiffTB = oDiffTB & " and (nvl(a0k06,0)+nvl(a0k07,0)) > (nvl(a0k17,0)+nvl(a0k18,0)) "
''   oDiffTB = oDiffTB & " GROUP BY A0K01,nvl(A0K06,0),nvl(A0K07,0) ) AA "
''   oDiffTB = oDiffTB & " where (nvl(a0k06,0)+nvl(a0k07,0)) > PAY GROUP BY a0k01 "
'
'   strSql = "delete R100123 where id='" & strUserNum & "'"
'   cnnConnection.Execute strSql, intI
'
'   strSql = " insert into R100123 (ID,R01,R02) " & _
'      " select '" & strUserNum & "' ID,a0k01 R01,nvl(a0k06,0)+nvl(a0k07,0)-nvl(PAY,0) R02 FROM(  " & _
'      "  SELECT A0K01,max(nvl(A0K06,0)) as a0k06,max(nvl(A0K07,0)) as a0k07" & _
'      " ,sum(nvl(a1u04,0))+sum(nvl(a1u07,0))-sum(nvl(a1u08,0))+sum(nvl(a1u05,0))+sum(nvl(a1u09,0))-sum(nvl(a1u10,0)) PAY " & _
'      " from acc0k0,ACC1U0 WHERE a0k02 <" & strSrvDate(2) & " AND (a0k09 is null or a0k09 = 0) " & stConAcc0k0 & _
'      " and (nvl(a0k06,0)+nvl(a0k07,0)) > (nvl(a0k17,0)+nvl(a0k18,0))  AND A0K01=A1U02(+) " & _
'      " GROUP BY A0K01) AA where a0k06+a0k07 > PAY"
'
'   cnnConnection.Execute strSql, intI
'
'   oDiffTB = "select R01 oKey,R02 diff FROM R100123 where ID='" & strUserNum & "'"
'
'   'Add by Morgan 2010/1/27 預定收款日資料
'   'Modify by Morgan 2010/12/7 已取消收文也要
'   stVTBw = "select rd01,substrb(max(rd02||(1000+rd03)||rd05),13) rd05" & _
'      " From R100123, caseprogress, ReceivablesDay" & _
'      " where id='" & strUserNum & "' and cp60(+)=R01" & _
'      " and rd01(+)=cp09 group by rd01"
'
''edit by nickc 2008/05/21 收文號改收據號碼
''edit by nickc 2008/05/29 修改踢除台灣
''Modify by Morgan 2010/1/27 抓最新預定收款日改用虛擬表格
''Modify by Morgan 2010/12/7 已取消收文也要
'   stCon = stCon & " union select s2.st06 as 所別,a0k22 as 管制人部門,'' as 管制人,a0k20 as 收文智權人員,'6' as 分類,' '||sqldatet(rd05) as 本所期限,'' as 法定期限,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,pa47 as 分所號,NVL(NVL(PA05,PA06),PA07) As 案件名稱,NVL(DECODE(PA09,'000',CPM03,CPM04),cp10) as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家,pa11 as 申請案號,a0k01 as 收文號,0 as 序號,'' as PKey " & _
'                  " from (" & stVTBw & ") BB,acc0k0,patent,caseprogress,nation,staff s2,customer,casepropertymap,(" & oDiffTB & ") difftb where BB.rd01=cp09(+) and cp60=a0k01 and (a0k09 is null or a0k09 = 0) and cp60=difftb.okey(+) and difftb.diff <> 0 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp01=cpm01(+) and cp10=cpm02(+) and pa09=na01(+) and a0k20= s2.st01(+) and substr(a0k03,1,8)=cu01(+) and substr(a0k03,9,1)=cu02(+) and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ") and pa09<>'000' " & stConST & stCon4
'   stCon = stCon & " union select s2.st06 as 所別,a0k22 as 管制人部門,'' as 管制人,a0k20 as 收文智權人員,'6' as 分類,' '||sqldatet(rd05) as 本所期限,'' as 法定期限,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,tm34 as 分所號,NVL(NVL(tm05,TM06),TM07) As 案件名稱,NVL(DECODE(tm10,'000',CPM03,CPM04),cp10) as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家,tm12 as 申請案號,a0k01 as 收文號,0 as 序號,'' as PKey " & _
'                  " from (" & stVTBw & ") BB,acc0k0,trademark,caseprogress,nation,staff s2,customer,casepropertymap,(" & oDiffTB & ") difftb where  BB.rd01=cp09(+) and cp60=a0k01 and (a0k09 is null or a0k09 = 0) and cp60=difftb.okey(+) and difftb.diff <> 0 and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp01=cpm01(+) and cp10=cpm02(+) and tm10=na01(+) and a0k20=s2.st01(+) and substr(a0k03,1,8)=cu01(+) and substr(a0k03,9,1)=cu02(+) and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ") and tm10<>'000' " & stConST & stCon4
'   stCon = stCon & " union select s2.st06 as 所別,a0k22 as 管制人部門,'' as 管制人,a0k20 as 收文智權人員,'6' as 分類,' '||sqldatet(rd05) as 本所期限,'' as 法定期限,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,sp28 as 分所號,NVL(NVL(SP05,SP06),SP07) As 案件名稱,NVL(DECODE(sp09,'000',CPM03,CPM04),cp10) as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家,sp11 as 申請案號,a0k01 as 收文號,0 as 序號,'' as PKey " & _
'                  " from (" & stVTBw & ") BB,acc0k0,servicepractice,caseprogress,nation,staff s2,customer,casepropertymap,(" & oDiffTB & ") difftb where BB.rd01=cp09(+) and cp60=a0k01 and (a0k09 is null or a0k09 = 0) and cp60=difftb.okey(+) and difftb.diff <> 0 and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and cp01=cpm01(+) and cp10=cpm02(+) and sp09=na01(+) and a0k20=s2.st01(+) and substr(a0k03,1,8)=cu01(+) and substr(a0k03,9,1)=cu02(+) and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 5) & ") and sp09<>'000' " & stConST & stCon4
'   stCon = stCon & " union select s2.st06 as 所別,a0k22 as 管制人部門,'' as 管制人,a0k20 as 收文智權人員,'6' as 分類,' '||sqldatet(rd05) as 本所期限,'' as 法定期限,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,lc16 as 分所號,NVL(NVL(LC05,LC06),LC07) As 案件名稱,NVL(DECODE(lc15,'000',CPM03,CPM04),cp10) as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家,'' as 申請案號,a0k01 as 收文號,0 as 序號,'' as PKey " & _
'                  " from (" & stVTBw & ") BB,acc0k0,lawcase,caseprogress,nation,staff s2,customer,casepropertymap,(" & oDiffTB & ") difftb where BB.rd01=cp09(+) and cp60=a0k01 and (a0k09 is null or a0k09 = 0) and cp60=difftb.okey(+) and difftb.diff <> 0 and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and cp01=cpm01(+) and cp10=cpm02(+) and lc15=na01(+) and a0k20=s2.st01(+) and substr(a0k03,1,8)=cu01(+) and substr(a0k03,9,1)=cu02(+) and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 3) & ") and lc15<>'000' " & stConST & stCon4
   
   'Add By Sindy 2014/6/12
   If Check2.Value = 1 Then '6.未收款
   '2014/6/12 END
      '目前有預定收款日的所有未收款收文資料
      stVTBw = "select rd01,substrb(max(rd02||(1000+rd03)||rd05),13) rd05" & _
         " from ReceivablesDay where rd06 is null group by rd01"
      '2011/8/25 modify by sonia 因收文即有cp79但開請款單者不會更新cp79故加cp60<'X'
      'Modified by Morgan 2011/11/22 考慮拆收據情形,收據號改用 getunpayno 函數抓未收款收據號
'      stCon = stCon & " union select s2.st06 as 所別,cp12 as 管制人部門,'' as 管制人,cp13 as 收文智權人員,'6' as 分類,'' as 約定期限,' '||sqldatet(rd05) as 本所期限,'' as 法定期限,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,pa47 as 分所號,NVL(NVL(PA05,PA06),PA07) As 案件名稱,NVL(DECODE(PA09,'000',CPM03,CPM04),cp10) as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家,pa11 as 申請案號,getunpayno(cp09) as 收文號,0 as 序號,'' as PKey " & _
'                     " from (" & stVTBw & ") BB,caseprogress,patent,nation,staff s2,customer,casepropertymap where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ") " & stCon8 & " and cp09(+)=BB.rd01 and cp79>0 and cp60 is not null and cp60<'X' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp01=cpm01(+) and cp10=cpm02(+) and pa09=na01(+) and s2.st01(+)=cp13 and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1) and pa09<>'000' " & stConST & stConP
      'Modified by Lydia 2017/06/21 針對超過預定收款日未收款之控制,取消申請國家<>'000'的限制
      'Modified by Lydia 2017/06/22 還原
      'Modified by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期(不限制已發文)
      'strSql = "insert into R100123_2" & _
              " select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'6' as 分類,'" & strUserNum & "',' '||sqldatet(rd05) as 本所期限,'' as 管制人,'' as 法定期限,'' as 約定期限,pa47 as 分所號,NVL(NVL(PA05,PA06),PA07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,pa26 as 申請人,pa09 as 申請國家,pa11 as 申請案號,0 as 序號,'' as PKey " & _
                " from (" & stVTBw & ") BB,caseprogress,patent,staff s2 where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ") " & stCon8 & " and cp09(+)=BB.rd01 and cp79>0 and cp60 is not null and cp60<'X' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and s2.st01(+)=cp13 and pa09<>'000' " & stConST & stConP
      'Modified by Lydia 2021/04/20 調整符合未收款之條件:(國內收據acc0k0)抓CP60 < 'X' + CP79>0(未收金額) 或 (國外請款acc1k0) CP60 < 'X' + Nvl(A1k29,'N')=未結清帳款
      'strSql = "insert into R100123_2" & _
                   " select cp01,cp02,cp03,cp04,cp09,管制人部門,cp13,分類,mid,本所期限,管制人,法定期限,約定期限,分所號, 案件名稱,案件性質,承辦人, 收文日 , 發文日, 申請人, 申請國家, 申請案號, 序號, pKey" & _
                   " from (select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'6' as 分類,'" & strUserNum & "' as mid,' '||sqldatet(cp06) as 本所期限,'' as 管制人,' '||sqldatet(cp07) as 法定期限,'' as 約定期限,pa47 as 分所號,NVL(NVL(PA05,PA06),PA07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,pa26 as 申請人,pa09 as 申請國家,pa11 as 申請案號,0 as 序號,'' as PKey,nvl(cu175,2) as cu175,cp60" & _
                   " from caseprogress,patent,staff s2,customer where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ") " & stCon8 & "  and cp79>0 and cp60 is not null and cp60<'X' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and s2.st01(+)=cp13 and pa09<>'000' " & stConST & stConP & _
                   " and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) ) aa,acc0k0,acc1k0" & _
                   " where cp60=a0k01(+) and cp60=a1k01(+) and ((nvl(a0k02,9999999)<(to_char(add_months(sysdate-1,cu175 * -1),'YYYYMMDD')-19110000) and a0k32 is null)" & _
                   " or nvl(a1k02,9999999)<(to_char(add_months(sysdate-1,cu175 * -1),'YYYYMMDD')-19110000))"
      'Modified by Lydia 2025/06/09 a0k32 is null 改用函數判斷：geta0k32type(a0k01)='1'
      strSql = "insert into R100123_2" & _
                   " select cp01,cp02,cp03,cp04,cp09,管制人部門,cp13,分類,mid,本所期限,管制人,法定期限,約定期限,分所號, 案件名稱,案件性質,承辦人, 收文日 , 發文日, 申請人, 申請國家, 申請案號, 序號, pKey" & _
                   " from (select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'6' as 分類,'" & strUserNum & "' as mid,' '||sqldatet(cp06) as 本所期限,'' as 管制人,' '||sqldatet(cp07) as 法定期限,'' as 約定期限,pa47 as 分所號,NVL(NVL(PA05,PA06),PA07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,pa26 as 申請人,pa09 as 申請國家,pa11 as 申請案號,0 as 序號,'' as PKey,nvl(cu175,2) as cu175,cp60,cp79" & _
                   " from caseprogress,patent,staff s2,customer where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ") " & stCon8 & "  and cp79>0 and cp60 is not null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and s2.st01(+)=cp13 and pa09<>'000' " & stConST & stConP & _
                   " and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) ) aa,acc0k0,acc1k0" & _
                   " where cp60=a0k01(+) and cp60=a1k01(+) and ( (Cp60 <'X' And Cp79>0 And Nvl(A0k02,9999999)<(To_Char(Add_Months(Sysdate-1,Cu175 * -1),'YYYYMMDD')-19110000) And geta0k32type(a0k01)='1') Or " & _
                   " (Cp60 >'X' And Nvl(A1k29,'N') <> 'Y' And Nvl(A1k02,9999999)<(To_Char(Add_Months(Sysdate-1,Cu175 * -1),'YYYYMMDD')-19110000)) )"
      cnnConnection.Execute strSql, intI
      
'      stCon = stCon & " union select s2.st06 as 所別,cp12 as 管制人部門,'' as 管制人,cp13 as 收文智權人員,'6' as 分類,'' as 約定期限,' '||sqldatet(rd05) as 本所期限,'' as 法定期限,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,tm34 as 分所號,NVL(NVL(tm05,TM06),TM07) As 案件名稱,NVL(DECODE(tm10,'000',CPM03,CPM04),cp10) as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家,tm12 as 申請案號,getunpayno(cp09) as 收文號,0 as 序號,'' as PKey " & _
'                     " from (" & stVTBw & ") BB,caseprogress,trademark,nation,staff s2,customer,casepropertymap where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ") " & stCon8 & " and cp09(+)=BB.rd01 and cp79>0 and cp60 is not null and cp60<'X' and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp01=cpm01(+) and cp10=cpm02(+) and tm10=na01(+) and s2.st01(+)=cp13 and cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9,1) and tm10<>'000' " & stConST & stConT
      'Modified by Lydia 2017/06/21 針對超過預定收款日未收款之控制,取消申請國家<>'000'的限制
      'Modified by Lydia 2017/06/22 還原
      'Modified by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
      'strSql = "insert into R100123_2" & _
              " select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'6' as 分類,'" & strUserNum & "',' '||sqldatet(rd05) as 本所期限,'' as 管制人,'' as 法定期限,'' as 約定期限,tm34 as 分所號,NVL(NVL(tm05,TM06),TM07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,tm23 as 申請人,tm10 as 申請國家,tm12 as 申請案號,0 as 序號,'' as PKey " & _
                " from (" & stVTBw & ") BB,caseprogress,trademark,staff s2 where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ") " & stCon8 & " and cp09(+)=BB.rd01 and cp79>0 and cp60 is not null and cp60<'X' and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and s2.st01(+)=cp13 and tm10<>'000' " & stConST & stConT
      'Modified by Lydia 2021/04/20 調整符合未收款之條件:(國內收據acc0k0)抓CP60 < 'X' + CP79>0(未收金額) 或 (國外請款acc1k0) CP60 < 'X' + Nvl(A1k29,'N')=未結清帳款
      'strSql = "insert into R100123_2" & _
                   " select cp01,cp02,cp03,cp04,cp09,管制人部門,cp13,分類,mid,本所期限,管制人,法定期限,約定期限,分所號, 案件名稱,案件性質,承辦人, 收文日 , 發文日, 申請人, 申請國家, 申請案號, 序號, pKey" & _
                   " from (select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'6' as 分類, '" & strUserNum & "' as mid,' '||sqldatet(cp06) as 本所期限,'' as 管制人,' '||sqldatet(cp07) as 法定期限,'' as 約定期限,tm34 as 分所號,NVL(NVL(tm05,tm06),tm07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,tm23 as 申請人,tm10 as 申請國家,tm12 as 申請案號,0 as 序號,'' as PKey,nvl(cu175,2) as cu175,cp60" & _
                   " from caseprogress,trademark,staff s2,customer where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ") " & stCon8 & "  and cp79>0 and cp60 is not null and cp60<'X' and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and s2.st01(+)=cp13 and tm10<>'000' " & stConST & stConT & _
                   " and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) ) aa,acc0k0,acc1k0" & _
                   " where cp60=a0k01(+) and cp60=a1k01(+) and ((nvl(a0k02,9999999)<(to_char(add_months(sysdate-1,cu175 * -1),'YYYYMMDD')-19110000) and a0k32 is null)" & _
                   " or nvl(a1k02,9999999)<(to_char(add_months(sysdate-1,cu175 * -1),'YYYYMMDD')-19110000))"
      'Modified by Lydia 2025/06/09 a0k32 is null 改用函數判斷：geta0k32type(a0k01)='1'
      strSql = "insert into R100123_2" & _
                   " select cp01,cp02,cp03,cp04,cp09,管制人部門,cp13,分類,mid,本所期限,管制人,法定期限,約定期限,分所號, 案件名稱,案件性質,承辦人, 收文日 , 發文日, 申請人, 申請國家, 申請案號, 序號, pKey" & _
                   " from (select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'6' as 分類, '" & strUserNum & "' as mid,' '||sqldatet(cp06) as 本所期限,'' as 管制人,' '||sqldatet(cp07) as 法定期限,'' as 約定期限,tm34 as 分所號,NVL(NVL(tm05,tm06),tm07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,tm23 as 申請人,tm10 as 申請國家,tm12 as 申請案號,0 as 序號,'' as PKey,nvl(cu175,2) as cu175,cp60,cp79" & _
                   " from caseprogress,trademark,staff s2,customer where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ") " & stCon8 & "  and cp79>0 and cp60 is not null and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and s2.st01(+)=cp13 and tm10<>'000' " & stConST & stConT & _
                   " and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) ) aa,acc0k0,acc1k0" & _
                   " where cp60=a0k01(+) and cp60=a1k01(+) and ( (Cp60 <'X' And Cp79>0 And Nvl(A0k02,9999999)<(To_Char(Add_Months(Sysdate-1,Cu175 * -1),'YYYYMMDD')-19110000) And geta0k32type(a0k01)='1') Or " & _
                   " (Cp60 >'X' And Nvl(A1k29,'N') <> 'Y' And Nvl(A1k02,9999999)<(To_Char(Add_Months(Sysdate-1,Cu175 * -1),'YYYYMMDD')-19110000)) )"
      cnnConnection.Execute strSql, intI
      
'      stCon = stCon & " union select s2.st06 as 所別,cp12 as 管制人部門,'' as 管制人,cp13 as 收文智權人員,'6' as 分類,'' as 約定期限,' '||sqldatet(rd05) as 本所期限,'' as 法定期限,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,sp28 as 分所號,NVL(NVL(SP05,SP06),SP07) As 案件名稱,NVL(DECODE(sp09,'000',CPM03,CPM04),cp10) as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家,sp11 as 申請案號,getunpayno(cp09) as 收文號,0 as 序號,'' as PKey " & _
'                     " from (" & stVTBw & ") BB,caseprogress,servicepractice,nation,staff s2,customer,casepropertymap where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 5) & ") " & stCon8 & " and cp09(+)=BB.rd01 and cp79>0 and cp60 is not null and cp60<'X' and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and cp01=cpm01(+) and cp10=cpm02(+) and sp09=na01(+) and s2.st01(+)=cp13 and cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9,1) and sp09<>'000' " & stConST & stConS
      'Modified by Lydia 2017/06/21 針對超過預定收款日未收款之控制,取消申請國家<>'000'的限制
      'Modified by Lydia 2017/06/22 還原
      'Modified by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
      'strSql = "insert into R100123_2" & _
              " select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'6' as 分類,'" & strUserNum & "',' '||sqldatet(rd05) as 本所期限,'' as 管制人,'' as 法定期限,'' as 約定期限,sp28 as 分所號,NVL(NVL(SP05,SP06),SP07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,sp08 as 申請人,sp09 as 申請國家,sp11 as 申請案號,0 as 序號,'' as PKey " & _
                " from (" & stVTBw & ") BB,caseprogress,servicepractice,staff s2 where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 5) & ") " & stCon8 & " and cp09(+)=BB.rd01 and cp79>0 and cp60 is not null and cp60<'X' and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and s2.st01(+)=cp13 and sp09<>'000' " & stConST & stConS
      'Modified by Lydia 2021/04/20 調整符合未收款之條件:(國內收據acc0k0)抓CP60 < 'X' + CP79>0(未收金額) 或 (國外請款acc1k0) CP60 < 'X' + Nvl(A1k29,'N')=未結清帳款
      'strSql = "insert into R100123_2" & _
                   " select cp01,cp02,cp03,cp04,cp09,管制人部門,cp13,分類,mid,本所期限,管制人,法定期限,約定期限,分所號, 案件名稱,案件性質,承辦人, 收文日 , 發文日, 申請人, 申請國家, 申請案號, 序號, pKey" & _
                   " from (select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'6' as 分類,'" & strUserNum & "' as mid,' '||sqldatet(cp06) as 本所期限,'' as 管制人,' '||sqldatet(cp07) as 法定期限,'' as 約定期限,sp28 as 分所號,NVL(NVL(sp05,sp06),sp07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,sp08 as 申請人,sp09 as 申請國家,sp11 as 申請案號,0 as 序號,'' as PKey,nvl(cu175,2) as cu175,cp60" & _
                   " from caseprogress,servicepractice,staff s2,customer where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 5) & ") " & stCon8 & "  and cp79>0 and cp60 is not null and cp60<'X' and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and s2.st01(+)=cp13 and sp09<>'000' " & stConST & stConS & _
                   " and substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+) ) aa,acc0k0,acc1k0" & _
                   " where cp60=a0k01(+) and cp60=a1k01(+) and ((nvl(a0k02,9999999)<(to_char(add_months(sysdate-1,cu175 * -1),'YYYYMMDD')-19110000) and a0k32 is null)" & _
                   " or nvl(a1k02,9999999)<(to_char(add_months(sysdate-1,cu175 * -1),'YYYYMMDD')-19110000))"
      'Modified by Lydia 2025/06/09 a0k32 is null 改用函數判斷：geta0k32type(a0k01)='1'
      strSql = "insert into R100123_2" & _
                   " select cp01,cp02,cp03,cp04,cp09,管制人部門,cp13,分類,mid,本所期限,管制人,法定期限,約定期限,分所號, 案件名稱,案件性質,承辦人, 收文日 , 發文日, 申請人, 申請國家, 申請案號, 序號, pKey" & _
                   " from (select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'6' as 分類,'" & strUserNum & "' as mid,' '||sqldatet(cp06) as 本所期限,'' as 管制人,' '||sqldatet(cp07) as 法定期限,'' as 約定期限,sp28 as 分所號,NVL(NVL(sp05,sp06),sp07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,sp08 as 申請人,sp09 as 申請國家,sp11 as 申請案號,0 as 序號,'' as PKey,nvl(cu175,2) as cu175,cp60,cp79" & _
                   " from caseprogress,servicepractice,staff s2,customer where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 5) & ") " & stCon8 & "  and cp79>0 and cp60 is not null and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and s2.st01(+)=cp13 and sp09<>'000' " & stConST & stConS & _
                   " and substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+) ) aa,acc0k0,acc1k0" & _
                   " where cp60=a0k01(+) and cp60=a1k01(+) and ( (Cp60 <'X' And Cp79>0 And Nvl(A0k02,9999999)<(To_Char(Add_Months(Sysdate-1,Cu175 * -1),'YYYYMMDD')-19110000) And geta0k32type(a0k01)='1') Or " & _
                   " (Cp60 >'X' And Nvl(A1k29,'N') <> 'Y' And Nvl(A1k02,9999999)<(To_Char(Add_Months(Sysdate-1,Cu175 * -1),'YYYYMMDD')-19110000)) )"
      cnnConnection.Execute strSql, intI
'      stCon = stCon & " union select s2.st06 as 所別,cp12 as 管制人部門,'' as 管制人,cp13 as 收文智權人員,'6' as 分類,'' as 約定期限,' '||sqldatet(rd05) as 本所期限,'' as 法定期限,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,lc16 as 分所號,NVL(NVL(LC05,LC06),LC07) As 案件名稱,NVL(DECODE(lc15,'000',CPM03,CPM04),cp10) as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家,'' as 申請案號,getunpayno(cp09) as 收文號,0 as 序號,'' as PKey " & _
'                     " from (" & stVTBw & ") BB,caseprogress,lawcase,nation,staff s2,customer,casepropertymap where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 3) & ") " & stCon8 & " and cp09(+)=BB.rd01 and cp79>0 and cp60 is not null and cp60<'X' and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and cp01=cpm01(+) and cp10=cpm02(+) and lc15=na01(+) and s2.st01(+)=cp13 and cu01(+)=substr(lc11,1,8) and cu02(+)=substr(lc11,9,1) and lc15<>'000' " & stConST & stConL
      'Modified by Lydia 2017/06/21 針對超過預定收款日未收款之控制,取消申請國家<>'000'的限制
      'Modified by Lydia 2017/06/22 還原
      'Modified by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
      'strSql = "insert into R100123_2" & _
              " select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'6' as 分類,'" & strUserNum & "',' '||sqldatet(rd05) as 本所期限,'' as 管制人,'' as 法定期限,'' as 約定期限,lc16 as 分所號,NVL(NVL(LC05,LC06),LC07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,lc11 as 申請人,lc15 as 申請國家,'' as 申請案號,0 as 序號,'' as PKey " & _
                " from (" & stVTBw & ") BB,caseprogress,lawcase,staff s2 where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 3) & ") " & stCon8 & " and cp09(+)=BB.rd01 and cp79>0 and cp60 is not null and cp60<'X' and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and s2.st01(+)=cp13 and lc15<>'000' " & stConST & stConL
      'Modified by Lydia 2021/04/20 調整符合未收款之條件:(國內收據acc0k0)抓CP60 < 'X' + CP79>0(未收金額) 或 (國外請款acc1k0) CP60 < 'X' + Nvl(A1k29,'N')=未結清帳款
      'strSql = "insert into R100123_2" & _
                   " select cp01,cp02,cp03,cp04,cp09,管制人部門,cp13,分類,mid,本所期限,管制人,法定期限,約定期限,分所號, 案件名稱,案件性質,承辦人, 收文日 , 發文日, 申請人, 申請國家, 申請案號, 序號, pKey" & _
                   " from (select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'6' as 分類,'" & strUserNum & "' as mid,' '||sqldatet(cp06) as 本所期限,'' as 管制人,' '||sqldatet(cp07) as 法定期限,'' as 約定期限,lc16 as 分所號,NVL(NVL(lc05,lc06),lc07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,lc11 as 申請人,lc15 as 申請國家,'' as 申請案號,0 as 序號,'' as PKey,nvl(cu175,2) as cu175,cp60" & _
                   " from caseprogress,lawcase,staff s2,customer where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 3) & ") " & stCon8 & "  and cp79>0 and cp60 is not null and cp60<'X' and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and s2.st01(+)=cp13 and lc09<>'000' " & stConST & stConL & _
                   " and substr(lc11,1,8)=cu01(+) and substr(lc11,9,1)=cu02(+) ) aa,acc0k0,acc1k0" & _
                   " where cp60=a0k01(+) and cp60=a1k01(+) and ((nvl(a0k02,9999999)<(to_char(add_months(sysdate-1,cu175 * -1),'YYYYMMDD')-19110000) and a0k32 is null)" & _
                   " or nvl(a1k02,9999999)<(to_char(add_months(sysdate-1,cu175 * -1),'YYYYMMDD')-19110000))"
      'Modified by Lydia 2025/06/09 a0k32 is null 改用函數判斷：geta0k32type(a0k01)='1'
      strSql = "insert into R100123_2" & _
                   " select cp01,cp02,cp03,cp04,cp09,管制人部門,cp13,分類,mid,本所期限,管制人,法定期限,約定期限,分所號, 案件名稱,案件性質,承辦人, 收文日 , 發文日, 申請人, 申請國家, 申請案號, 序號, pKey" & _
                   " from (select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'6' as 分類,'" & strUserNum & "' as mid,' '||sqldatet(cp06) as 本所期限,'' as 管制人,' '||sqldatet(cp07) as 法定期限,'' as 約定期限,lc16 as 分所號,NVL(NVL(lc05,lc06),lc07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,lc11 as 申請人,lc15 as 申請國家,'' as 申請案號,0 as 序號,'' as PKey,nvl(cu175,2) as cu175,cp60,cp79" & _
                   " from caseprogress,lawcase,staff s2,customer where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 3) & ") " & stCon8 & "  and cp79>0 and cp60 is not null and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and s2.st01(+)=cp13 and lc09<>'000' " & stConST & stConL & _
                   " and substr(lc11,1,8)=cu01(+) and substr(lc11,9,1)=cu02(+) ) aa,acc0k0,acc1k0" & _
                   " where cp60=a0k01(+) and cp60=a1k01(+) and ( (Cp60 <'X' And Cp79>0 And Nvl(A0k02,9999999)<(To_Char(Add_Months(Sysdate-1,Cu175 * -1),'YYYYMMDD')-19110000) And geta0k32type(a0k01)='1') Or " & _
                   " (Cp60 >'X' And Nvl(A1k29,'N') <> 'Y' And Nvl(A1k02,9999999)<(To_Char(Add_Months(Sysdate-1,Cu175 * -1),'YYYYMMDD')-19110000)) )"
      cnnConnection.Execute strSql, intI
'end 2011/8/15
   End If
   
   'Add By Sindy 2014/6/12
   If Check4.Value = 1 Then '1.未發文
   '2014/6/12 END
      'Add By Sindy 2012/6/4 T、FCT台灣商標爭議案逾承辦期限、逾指定會稿日
'      stCon = stCon & " union select s2.st06 as 所別,cp12 as 管制人部門,cp14 as 管制人,cp13 as 收文智權人員,'1' as 分類,'' as 約定期限,decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,tm34 as 分所號,NVL(NVL(tm05,TM06),TM07) As 案件名稱,NVL(DECODE(tm10,'000',CPM03,CPM04),cp10) as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家,tm12 as 申請案號,cp09 as 收文號,0 as 序號,'' as PKey " & _
'                     " from trademark,caseprogress,nation,staff s2,customer,casepropertymap,t102inform,EngineerProgress " & _
'                     " where CP05>=20120601 and cp01 in ('T','FCT') " & _
'                     " and cp27 is null and cp57 is null AND CP10 in (" & TMdebate & ") " & _
'                     " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & _
'                     " and TM10='000' and tm29 is null " & _
'                     " and cp01=cpm01(+) and cp10=cpm02(+) " & _
'                     " and tm10=na01(+) and cp13=s2.st01(+) " & _
'                     " and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) " & stConST & stCon2 & stConT & _
'                     " and CP09=EP02(+) and cp09=ti02(+) " & _
'                     " and ((CP48<" & strSrvDate(1) & " and CP48 is not null) or (EP28<" & strSrvDate(1) & " and EP28 is not null))"
      '2012/6/4 End
      'Modified by Lydia 2016/12/30 +排除D類收文 cp27 is null and cp57 is null => nvl(cp27,0)=0 and nvl(cp57,0)=0 and substr(cp09,1,1)<>'D'
      'Modified by Lydia 2019/04/08 nvl(cp27,0)=0 and nvl(cp57,0)=0=> cp158=0 and cp159=0
      'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
      strSql = "insert into R100123_2" & _
              " select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'1' as 分類,'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,cp14 as 管制人,sqldatet(cp07) as 法定期限,'' as 約定期限,tm34 as 分所號,NVL(NVL(tm05,TM06),TM07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,tm23 as 申請人,tm10 as 申請國家,tm12 as 申請案號,0 as 序號,'' as PKey " & _
                " from trademark,caseprogress,staff s2,t102inform,EngineerProgress " & _
                  " where CP05>=20120601 and cp01 in ('T','FCT') " & _
                  " and cp158=0 and cp159=0 and substr(cp09,1,1)<>'D' AND CP10 in (" & TMdebate & ") And Not (cp01='FCT' And InStr(" & FCT_NotTMdebate & ", cp10) > 0) " & _
                  " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & _
                  " and TM10='000' and tm29 is null " & _
                  " and cp13=s2.st01(+) " & stConST & stCon2 & stConT & _
                  " and CP09=EP02(+) and cp09=ti02(+) " & _
                  " and ((CP48<" & strSrvDate(1) & " and CP48 is not null) or (EP28<" & strSrvDate(1) & " and EP28 is not null))"
      cnnConnection.Execute strSql, intI
      'Added by Lydia 2018/12/10 +T台灣案非爭議案
      If strSrvDate(1) >= T案收文齊備啟用日 Then
            'Modified by Lydia 2019/04/08 nvl(cp27,0)=0 and nvl(cp57,0)=0=> cp158=0 and cp159=0
            strSql = "insert into R100123_2" & _
                    " select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'1' as 分類,'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,cp14 as 管制人,sqldatet(cp07) as 法定期限,'' as 約定期限,tm34 as 分所號,NVL(NVL(tm05,TM06),TM07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,tm23 as 申請人,tm10 as 申請國家,tm12 as 申請案號,0 as 序號,'' as PKey " & _
                      " from trademark,caseprogress,staff s2,t102inform,EngineerProgress " & _
                        " where cp01 ='T' and cp05>=" & T案收文齊備啟用日 & _
                        " and cp158=0 and cp159=0 and substr(cp09,1,1)='A' AND CP10 not in (" & TMdebate & ") " & _
                        " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & _
                        " and TM10='000' and tm29 is null " & _
                        " and cp13=s2.st01(+) " & stConST & stCon2 & stConT & _
                        " and CP09=EP02(+) and cp09=ti02(+) " & _
                        " and ((CP48<" & strSrvDate(1) & " and CP48 is not null) or (EP28<" & strSrvDate(1) & " and EP28 is not null))"
            cnnConnection.Execute strSql, intI
      End If
      'end 2018/12/10
   End If
   
'   'Add By Sindy 2014/6/24 未處理:只顯示已通知
'   If Check6.Value = 1 Then
'      '1913.通知期限
'      'Modify By Sindy 2014/7/15 增加控管只有年費(605,606,607)及實體審查416才需檢查是否有1913通知期限,其他均顯示出來
'      'Modify By Sindy 2014/12/15 and rcp01='P' ==> and rcp01 in('P','CFP')
'      strSql = "delete from R100123_2 where id='" & strUserNum & "' and rkind='2' and rcp01 in('P','CFP') and rCP10 in ('605','606','607','416')" & _
'                 " and not exists (select cp09 from caseprogress where cp43=rcp09 and cp10='1913')"
'      cnnConnection.Execute strSql, intI
'   End If
'   '2014/6/24 END
   
   'Add By Sindy 2014/10/22 未處理:C類NP01->CP09若無CP27時,未處理則以未通知顯示 ex:CFP-26936
   'Modified by Lydia 2019/04/08 nvl(cp27,0)=0 => cp158=0
   strSql = "UPDATE R100123_2" & _
            " set rkind='5'" & _
            " where id='" & strUserNum & "' and rkind='2'" & _
            " and rcp09 in(select rcp09 from R100123_2,caseprogress where id='" & strUserNum & "' and rkind='2'" & _
                            " and rcp09=cp09 and cp158=0 and substr(rcp09,1,1)='C')"
   cnnConnection.Execute strSql, intI
   '2014/10/22 END
   
'Add By Sindy 2015/9/18 10.結案中 : 原為 未處理,未函知 的情形下,可能已填寫結案單, 但仍在結案流程中(程序尚未處理)
   '                                   請將此類資料的事件請改為 結案中
   '                                   結案中, 預設是否勾選與 未處理,未函知 相同
   'Modify By Sindy 2017/7/25 8碼為結案電子表單編號 ==>  and length(np24)=8
   'Modify by Amy 2020/05/18 T/TF延展、續展、第二期註冊費,若已結案且 未勾「結案中」則不顯示
   If Check7.Enabled = True Then
      'Modify By Sindy 2020/5/20
      strSql = "UPDATE R100123_2" & _
               " set rkind='10'" & _
               " where id='" & strUserNum & "' and rkind in('2','9')" & _
               " and exists (select np01,np22,np06,np24 from nextprogress where np01=RCP09 and np22=RNP22" & _
                           " and np06 is null and np24 is not null and length(np24)=8)"
      cnnConnection.Execute strSql, intI
'      If Check7.Value = 1 Then
'            strSql = "UPDATE R100123_2" & _
'                     " set rkind='10'" & _
'                     " where id='" & strUserNum & "' and rkind in('2','9')" & _
'                     " and exists (select np01,np22,np06,np24 from nextprogress where np01=RCP09 and np22=RNP22" & _
'                                 " and np06 is null and np24 is not null and length(np24)=8)"
'            cnnConnection.Execute strSql, intI
'      Else
      If Check7.Value = 0 Then
         strSql = "Delete R100123_2" & _
                    " Where id='" & strUserNum & "' and rkind='10'"
         cnnConnection.Execute strSql, intI
      '2020/5/20 END
         strSql = "Delete R100123_2" & _
                    " Where id='" & strUserNum & "' And RCP01||RCP10 in ('T102','T716','TF102','TF716') And InStr(rcp06,'*')>0"
         cnnConnection.Execute strSql, intI
      End If
   End If
   'end 2020/05/18
   '2015/9/18 END
   
   'Mark by Amy 2020/03/26 只開放電腦中心
   'cmdOK(8).Enabled = False   'Add by Amy 2020/03/25 +資料寄信箱鈕
   
   'edit by nickc 2005/09/22 取消智權人員的排序
   'stCon = "select '' as V,decode(AA.所別,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 管制人部門,s3.st02 as 管制人,s1.st02 as 收文智權人員,decode(AA.分類,'1','未發文','2','未處理','3','行事曆','4','未回覆','5','未通知','6','未收款','') as 事件分類,AA.本所期限,AA.法定期限,AA.本所案號,AA.分所號,AA.案件名稱,AA.案件性質,s2.st02 as 承辦人,AA.收文日,AA.發文日,AA.申請人,AA.申請國家,AA.申請案號,AA.收文號,AA.序號,AA.PKey from (" & stCon & ") AA,acc090,staff S1,Staff S2,staff S3 where AA.管制人部門=a0901(+) and AA.收文智權人員=s1.st01(+) and AA.承辦人=s2.st01(+) and AA.管制人=s3.st01(+)  order by AA.所別,AA.管制人部門,AA.管制人,AA.分類,substr('0'||AA.本所期限,-9,3),AA.本所案號 " 'substr(AA.本所期限,2)
   'Modify By Sindy 2010/8/25
   'stCon = "select '' as V,decode(AA.所別,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 管制人部門,s3.st02 as 管制人,s1.st02 as 收文智權人員,decode(AA.分類,'1','未發文','2','未處理','3','行事曆','4','未回覆','5','未通知','6','未收款','') as 事件分類,AA.本所期限,AA.法定期限,AA.本所案號,AA.分所號,AA.案件名稱,AA.案件性質,s2.st02 as 承辦人,AA.收文日,AA.發文日,AA.申請人,AA.申請國家,AA.申請案號,AA.收文號,AA.序號,AA.PKey from (" & stCon & ") AA,acc090,staff S1,Staff S2,staff S3 where AA.管制人部門=a0901(+) and AA.收文智權人員=s1.st01(+) and AA.承辦人=s2.st01(+) and AA.管制人=s3.st01(+)  order by AA.所別,AA.管制人部門,AA.管制人,AA.分類,decode(substr(ltrim(AA.本所期限),1,1),'*',substr('0'||substr(ltrim(AA.本所期限),2,length(ltrim(AA.本所期限))),-9,9),substr('0'||ltrim(AA.本所期限),-9,9)),AA.本所案號 "
   'Modify By Sindy 2011/6/20 未續簽&未回執
   'If Left(Trim(UCase(stCon)), 5) = UCase("union") Then stCon = Mid(Trim(stCon), 6) 'Add By Sindy 2014/6/12
   'stCon = "select '' as V,decode(AA.所別,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 管制人部門,s3.st02 as 管制人,s1.st02 as 收文智權人員,decode(substr(AA.分類,1,1),'1','未發文','2','未處理','3','行事曆','4','未回覆','5','未通知','6','未收款','7','未續簽','8','未回執','') as 事件分類,AA.約定期限,AA.本所期限,AA.法定期限,AA.本所案號,AA.分所號,AA.案件名稱,AA.案件性質,s2.st02 as 承辦人,AA.收文日,AA.發文日,AA.申請人,AA.申請國家,AA.申請案號,AA.收文號,AA.序號,AA.PKey from (" & stCon & ") AA,acc090,staff S1,Staff S2,staff S3 where AA.管制人部門=a0901(+) and AA.收文智權人員=s1.st01(+) and AA.承辦人=s2.st01(+) and AA.管制人=s3.st01(+)  order by AA.所別,AA.管制人部門,AA.管制人,AA.分類,NVL(AA.約定期限,decode(substr(ltrim(AA.本所期限),1,1),'*',substr('0'||substr(ltrim(AA.本所期限),2,length(ltrim(AA.本所期限))),-9,9),substr('0'||ltrim(AA.本所期限),-9,9))),AA.本所案號 "
   'Modify By Sindy 2015/3/2 +,'9','未函知'
   'Modify by Amy 2020/05/18 本所期限+Replace
   'Modify by Amy 2020/05/20 內商延展結案=本所期限有*號,狀態顯示結案中
   'Modify by Sindy 2020/7/30 P-123766-0-00; remp(A4099) 改為 rcp13(A4023); 讓智權人員的資料排在前面點; 因當時文雄該筆未發文,在倒數第3筆
   '  " order by s1.st06,rcp12,remp,rkind,NVL(rnp23,decode(substr(ltrim(rcp06),1,1),'*',substr('0'||substr(ltrim(rcp06),2,length(ltrim(rcp06))),-9,9),substr('0'||ltrim(rcp06),-9,9))),rcp01,rcp02,rcp03,rcp04"
   '  " order by s1.st06,rcp12,rcp13,rkind,NVL(rnp23,decode(substr(ltrim(rcp06),1,1),'*',substr('0'||substr(ltrim(rcp06),2,length(ltrim(rcp06))),-9,9),substr('0'||ltrim(rcp06),-9,9))),rcp01,rcp02,rcp03,rcp04"
   stCon = "select '' as V,decode(s1.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 管制人部門,s3.st02 as 管制人,s1.st02 as 收文智權人員,Decode(Substr(rcp06,1,1),'*','結案中',decode(rkind,'1','未發文','2','未處理','3','行事曆','4','未回覆','5','未通知','6','未收款','7','未續簽','8','未回執','9','未函知','10','結案中','')) as 事件分類,rnp23 as 約定期限,Replace(rcp06,'*','') as 本所期限,rcp07 as 法定期限,decode(rcp01,'','',rcp01||'-'||rcp02||'-'||rcp03||'-'||rcp04) as 本所案號,rsubno as 分所號,rcasename as 案件名稱,NVL(DECODE(rnation,'000',CPM03,CPM04),rcp10) as 案件性質,s2.st02 as 承辦人,rcp05 as 收文日,rcp27 as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家,rcaseno as 申請案號,rcp09 as 收文號,rnp22 as 序號,rpkey as PKey" & _
            " from (select distinct * from R100123_2 where id='" & strUserNum & "'),acc090,staff S1,Staff S2,staff S3,nation,customer,casepropertymap" & _
            " where rcp01=cpm01(+) and rcp10=cpm02(+)" & _
            " and rnation=na01(+)" & _
            " and substr(rappid,1,8)=cu01(+) and substr(rappid,9,1)=cu02(+)" & _
            " and rcp12=a0901(+) and rcp13=s1.st01(+) and rcp14=s2.st01(+) and remp=s3.st01(+)" & _
            " order by s1.st06,rcp12,rcp13,rkind,NVL(rnp23,decode(substr(ltrim(rcp06),1,1),'*',substr('0'||substr(ltrim(rcp06),2,length(ltrim(rcp06))),-9,9),substr('0'||ltrim(rcp06),-9,9))),rcp01,rcp02,rcp03,rcp04"
   CheckOC3
   SetDataListWidth
   grdDataList.Rows = 2
   grdDataList.Clear
   'SetDataListWidth 'Mark by Lydia 2021/06/01 當有資料時,下方會重設
   With AdoRecordSet3
      .CursorLocation = adUseClient
      'Debug.Print Timer
      .Open stCon, cnnConnection, adOpenStatic, adLockReadOnly
      'Debug.Print Timer
      LblCntTime.Caption = LblCntTime.Caption & " ~ " & Format(ServerTime, "##:##:##") & " 共 " & .RecordCount & " 筆" 'Add By Sindy 2014/6/12
      If .RecordCount > 0 Then
         'Modified by Morgan 2012/8/21 +約定期限欄位(6),>=6以後的索引+1
         Set grdDataList.Recordset = AdoRecordSet3.Clone
         SetDataListWidth
         'add by nickc 2005/08/18 當天變淺紅
         grdDataList.Visible = False
         For i = 1 To grdDataList.Rows - 1
            grdDataList.row = i
            ' 相關案件性質  2012/7/18 ADD BY SONIA
            If .Fields("事件分類") = "未回覆" Then
               'Modified by Lydia 2021/05/10 改用變數
               'grdDataList.TextMatrix(i, 12) = grdDataList.TextMatrix(i, 12) & PUB_GetNextCasePropertyName(grdDataList.TextMatrix(i, 19), grdDataList.TextMatrix(i, 20), "1")
               grdDataList.TextMatrix(i, colCPM) = grdDataList.TextMatrix(i, colCPM) & PUB_GetNextCasePropertyName(grdDataList.TextMatrix(i, colNP01), grdDataList.TextMatrix(i, colNP22), "1")
            End If
            '2012/7/18 END
            'Modified by Lydia 2021/05/10 改用變數
            'grdDataList.col = 7 '本所期限
            grdDataList.col = colCP06 '本所期限
            'Add By Sindy 2012/5/31 檢查本所案號是否有轉案至他所,若有,則在本所期限前加※符號
            Dim strCP06 As String
            'Modified by Lydia 2021/05/10 改用變數
            'grdDataList.TextMatrix(i, 7) = PUB_GetCP10ValueAttachText(grdDataList.TextMatrix(i, 9), "728", "※", grdDataList.TextMatrix(i, 7))
            grdDataList.TextMatrix(i, colCP06) = PUB_GetCP10ValueAttachText(grdDataList.TextMatrix(i, colCaseNo), "728", "※", grdDataList.TextMatrix(i, colCP06))
            strCP06 = Trim(grdDataList.Text)
            If Left(strCP06, 1) = "※" Then
               strCP06 = Mid(strCP06, 2)
            End If
            If Left(strCP06, 1) = "*" Then
               strCP06 = Mid(strCP06, 2)
            End If
            Call recovercolor(i) 'Modify By Sindy 2014/6/25 Sindy 還原顏色
'            'If Mid(grdDataList.Text, 2) = ChangeTStringToTDateString(strSrvDate(2)) Then
'            If strCP06 = ChangeTStringToTDateString(strSrvDate(2)) Then
'            '2012/5/31 End
'               For j = 0 To GrdDataList.Cols - 1
'                     GrdDataList.col = j
'                     GrdDataList.CellBackColor = &H8080FF
'               Next j
'            'add by nickc 2007/01/23
'            Else
'               GrdDataList.col = 9 '本所案號
'               If UCase(Mid(GrdDataList.Text, 1, 3)) = "CFT" Then
'                  GrdDataList.col = 12
'                  If Trim(GrdDataList.Text) = "延展" Or Trim(GrdDataList.Text) = "使用宣誓" Then
'                     For j = 0 To GrdDataList.Cols - 1
'                        GrdDataList.col = j
'                        GrdDataList.CellBackColor = &HC000&
'                     Next j
'                  End If
'               End If
'            End If
         Next i
         grdDataList.Visible = True
         'Mark by Amy 2020/03/26
         'Add by Amy 2020/03/25 同業務區且智權人員 為下拉選單且為空值或 智權人員 非下拉選單且未輸入值 或電腦中心,顯示 資料寄信箱鈕
'         If (txtSalesArea = txtSalesArea1 And ((Combo3.Visible = True And Trim(Combo3) = "") Or (Combo3.Visible = False And Trim(txtSales) = ""))) _
'          Or Pub_StrUserSt03 = "M51" Then
'            cmdOK(8).Enabled = True
'         End If
         'end 2020/03/26
      Else
         If bolShowMsgBox = True Then
            MsgBox "無符合資料！", vbInformation
         End If
      End If
   End With
   
   doQuery = True
   If bolChgSystemkind = True Then systemkind = strOldSystemkind
   
   'Add By Sindy 2020/7/28
   If txtSales.Visible = True Then 'Added by Lydia 2021/05/20 排除隱藏
        If txtSales.Enabled = True Then
           txtSales_GotFocus
           txtSales.SetFocus
        End If
   End If 'Added by Lydia 2021/05/20
   '2020/7/28 END
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   'Resume
End Function

Private Sub Check1_Click()
   If Check1.Value = 1 Then
      Check5.Enabled = True
      Check6.Enabled = True
      Check7.Enabled = True: Check7.Value = 1 'Add By Sindy 2015/9/17
   Else
      Check5.Enabled = False
      Check6.Enabled = False
      Check5.Value = 0
      Check6.Value = 0
      Check7.Enabled = False: Check7.Value = 0 'Add By Sindy 2015/9/17
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdok_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Public Sub cmdSearch_Click()
Dim bolCancel As Boolean 'Add By Sindy 2016/5/5
   'add by nickc 2006/11/21
'   If Mid(txtSalesArea, 1, 1) <> "S" And Mid(txtSalesArea1, 1, 1) <> "S" Then
'       chkNP.Enabled = False
'       chkNP.Value = 0
'   Else
'       chkNP.Enabled = True
'   End If
   'Add By Sindy 2016/5/5
   If Combo3.Visible = True Then
      bolCancel = False
      Call Combo3_Validate(bolCancel)
      If bolCancel = True Then
         Exit Sub
      End If
   End If
   '2016/5/5 END
   
   Screen.MousePointer = vbHourglass
   grdDataList.MousePointer = flexHourglass
'   bolSelData = False
   If ConstrainCheck = True Then
      'SetDataListWidth 'Mark by Lydia 2021/06/01 doQuery有設定
      bolShowMsgBox = True
      Call doQuery
   End If
   grdDataList.MousePointer = flexDefault
   Screen.MousePointer = vbDefault
   
   'Mark by Amy 2020/03/25 不需使用-秀玲
   '2011/8/9 ADD BY SONIA
'   If txtSales = "68006" Or txtSales = "68096" Or strUserNum = "68006" Then
'      cmdOK(3).Visible = True
'   Else
'      cmdOK(3).Visible = False
'   End If
   '2011/8/9 END
   
   cmdOK(7).Enabled = False 'Added by Morgan 2016/3/2
End Sub

'Add By Sindy 2014/6/9
Private Sub CountMonthToDay()
   If Combo2.Text <> "" Then
      If Val(Combo2.Text) > 0 Then
         txtCloseDate(0) = ChangeWDateStringToTString(DateAdd("d", -2, ChangeWStringToWDateString(strSrvDate(1))))
         txtCloseDate(1).Text = Val(Format(DateAdd("M", Val(Combo2.Text), ChangeTStringToWDateString(txtCloseDate(0))), "YYYYMMDD")) - 19110000
      ElseIf Val(Combo2.Text) < 0 Then
         txtCloseDate(1) = ChangeWDateStringToTString(DateAdd("d", -2, ChangeWStringToWDateString(strSrvDate(1))))
         txtCloseDate(0).Text = Val(Format(DateAdd("M", Val(Combo2.Text), ChangeTStringToWDateString(txtCloseDate(1))), "YYYYMMDD")) - 19110000
      End If
      txtCloseDate(0).Tag = txtCloseDate(0).Text
      txtCloseDate(1).Tag = txtCloseDate(1).Text
   End If
End Sub
Private Sub Combo2_Change()
   Call CountMonthToDay
End Sub
Private Sub Combo2_Click()
   Call CountMonthToDay
End Sub
'2014/6/9 END

'Add By Sindy 2016/5/4
'Modified by Lydia 2021/05/10 改成Form 2.0
'Private Sub Combo3_KeyPress(KeyAscii As Integer)
Private Sub Combo3_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub Combo3_LostFocus()
   If Trim(Combo3) <> "" And Trim(Combo3) <> "全部" Then
      'Combo3 = Trim(Left(Combo3, 6)) & " " & GetPrjSalesNM(Trim(Left(Combo3, 6)))
      arrID = Split(Combo3, " ")
      txtSales = arrID(0)
      Combo3 = txtSales & " " & GetPrjSalesNM(txtSales)
   'Add Sindy 2019/4/16
   ElseIf Trim(Combo3) <> "全部" Then
      txtSales = ""
   '2019/4/16 END
   End If
End Sub
Private Sub Combo3_Validate(Cancel As Boolean)
Dim strEmp As String
Dim stTmp As String 'Add by Amy 2020/03/25
   
   'modify by sonia 2016/6/7 因有S29不足五碼故改寫法
   'If Combo3 <> "" And Trim(Combo3) <> "全部" Then
   '   strEmp = GetStaffName(Trim(Left(Combo3, 6)))
   '   If strEmp = "" Then
   '      MsgBox "智權人員輸入錯誤！", vbCritical
   '      Combo3.SetFocus
   '      Cancel = True
   '   End If
   '   txtSales = Trim(Left(Combo3, 6))
   '   txtSales = Combo3
   '
   '   lblSalesName.Caption = strEmp
   '   Combo3 = Trim(Left(Combo3, 6)) & " " & GetPrjSalesNM(Trim(Left(Combo3, 6)))
   'End If
   If Combo3 <> "" And Trim(Combo3) <> "全部" Then
      'Add by Amy 2020/03/25 只能輸入下拉選單中已有的人員
      stTmp = Combo3
      '直接輸員編未串名字會錯
      If InStr(stTmp, " ") > 0 Then
        stTmp = Mid(stTmp, 1, Val(InStr(stTmp, " ")) - 1)
      Else
        stTmp = Combo3
      End If
      'Modify By Sindy 2020/6/15 Mark
'      If InStr(m_strListPer, stTmp) = 0 And stTmp <> strUserNum And Pub_StrUserSt03 <> "M51" Then
'         MsgBox "不可輸入下拉選單以外的人員！"
'         Cancel = True
'         Combo3.SetFocus
'         Exit Sub
'      End If
      'end 2020/03/25
      arrID = Split(Combo3, " ")
      txtSales = arrID(0)
      lblSalesName.Caption = GetStaffName(txtSales, True)
      If lblSalesName.Caption = "" Then
         MsgBox "智權人員輸入錯誤！", vbCritical
         Combo3.SetFocus
         Cancel = True
      End If
      Combo3 = txtSales & " " & GetPrjSalesNM(txtSales)
   'Modify By Sindy 2024/8/5 mark; 因 txtSales_Validate 會檢查相關的權限
   Else
      txtSales = ""
'   'Modify by Amy 2023/05/09 +st05
'   ElseIf Combo3 = MsgText(601) And stST05 <> "00" And stST05 <> "01" And stST05 <> "08" Then
'        'Add by Amy 2020/03/25 下拉選單無區主管智權人員不可為空
'        'Modify By Sindy 2020/7/14
'        'If bolAreaMan = False And Pub_StrUserSt03 <> "M51" Then
'        'Modify By Sindy 2023/9/21 開放杜協理權限 + And Not (Mid(txtSalesArea, 1, 1) = "S" And Mid(txtSalesArea1, 1, 1) = "S" And InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) > 0)
'        If (GetDeptMan(txtSalesArea) <> strUserNum Or GetDeptMan(txtSalesArea1) <> strUserNum) _
'            And Pub_StrUserSt03 <> "M51" _
'            And Not (Mid(txtSalesArea, 1, 1) = "S" And Mid(txtSalesArea1, 1, 1) = "S" And InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) > 0) Then
'        '2020/7/14 END
'           MsgBox "非區主管職代智權人員不可空白！"
'           Cancel = True
'           Combo3.SetFocus
'           Exit Sub
'        End If
'        'end 2020/03/25
   '2024/8/5 END
   End If
   'end 2016/6/7
End Sub
'2016/5/4 END

Private Sub Form_Load()
'edit by nickc 2008/01/18 改成全域變數
'Dim stST05 As String, stST15 As String
Dim nFrm As Form 'Added by Lydia 2021/05/10

   'Modify by Amy 2020/03/25 從下面搬上來
   'stST15 = PUB_GetStaffST15(strUserNum, 1)
   stST05 = PUB_GetST05(strUserNum)
   'bolAreaMan = False
   'Modify by Amy 2020/03/26 只開放電腦中心用(權限判斷太複雜)
   'cmdOK(8).Enabled = False
   cmdOK(8).Visible = False
   If Pub_StrUserSt03 = "M51" Then
     cmdOK(8).Visible = True
   End If
   'end 2020/03/25
   
   'Added by Lydia 2021/05/10先判斷表單是否存在
   Me.cmdOK(3).Visible = False
   If strSrvDate(1) >= m_NewStartDate Then
      Set nFrm = Forms(0).GetForm("frm100123_2")
      If Not nFrm Is Nothing Then
          Me.cmdOK(3).Visible = True
      End If
   End If
   'end 2021/05/10
                     
   Combo1.Clear
   Combo1.AddItem "紅色：當天期限或逾期"
   Combo1.AddItem "綠色：外商延展、使用宣誓法定期限"
   'Mark by Amy 2020/05/18
   'Combo1.AddItem "( * )：本所期限前加 * 代表發過結案通知"
   Combo1.AddItem "(※)：本所期限前加※代表轉案至他所"
   Combo1.ListIndex = 0
   
   MoveFormToCenter Me
   bolShowMsgBox = False
   
   If pub_CallNextForm = True Then 'APP開啟時,自動Run
      '系統日前後2天,共5日
      txtCloseDate(0) = ChangeWDateStringToTString(DateAdd("d", -2, ChangeWStringToWDateString(strSrvDate(1)))) 'strSrvDate(2)
      txtCloseDate(1) = ChangeWDateStringToTString(DateAdd("d", 2, ChangeWStringToWDateString(strSrvDate(1))))  'strSrvDate(2)
'      txtCloseDate(0).Tag = txtCloseDate(0).Text
'      txtCloseDate(1).Tag = txtCloseDate(1).Text
   Else
      Combo2.ListIndex = 7 '3個月
   End If
   
   'Modify By Sindy 2020/7/28 設定員編,部門,所別權限
   'Modify By Sindy 2023/5/16 + txtCloseDate(0)
   Call PUB_SetFormSaleDept(strUserNum, txtZone, txtSalesArea, txtSalesArea1, txtSales, bolSpecMan, strSpecCode _
         , , , , , , True, txtCloseDate(0))
   
   'Add By Sindy 2016/5/3
   '檢查當時是否需要為他人職代
   Combo3.Clear
   'Add By Sindy 2023/5/16
   If txtSales <> strUserNum And txtSales <> "" Then
      Combo3.AddItem txtSales & " " & GetPrjSalesNM(txtSales)
   End If
   '2023/5/16 END
   Combo3.AddItem strUserNum & " " & strUserName
   Call Pub_SetForOthersEmpCombo(strUserNum, Combo3, False, m_strListPer)
   If m_strListPer = "" Then
      Combo3.Visible = False
   Else
      'Add by Amy 2020/03/25 判斷下拉選單是否有區主管
'      If InStr(m_strListPer, GetDeptMan(stST15)) > 0 Then
'         bolAreaMan = True
'      End If
      Combo3.Visible = True
      Combo3.ListIndex = 0
      'Added by Lydia 2021/05/20 Form 2.0物件無法覆蓋Form 1.0
      txtSales.Visible = False
      lblSalesName.Visible = False
   End If
   '2016/5/3 END
   
   SetDataListWidth
'   bolSelData = False
   'Modify by Morgan 2011/4/21 從 Unload 移來(因畫面沒離開時沒寫Log會造成逾時重新登入後重複執行)
   PUB_AddExcuteLog Me.Name 'Add By Sindy 2011/3/17
   
   'Add By Sindy 2014/6/9
   If pub_CallNextForm = True Then
      Check1.Value = 1 '未處理、未函知
      Check2.Value = 1 '未收款
      Check3.Value = 1 '未通知
      Check4.Value = 1 '未發文
      chkNP.Value = 0 '未回覆
      Check7.Value = 1 'Add By Sindy 2015/9/17 結案中
   Else
      Check1.Value = 1
      Check2.Value = 0
      Check3.Value = 0
      Check4.Value = 1
      chkNP.Value = 0
      Check7.Value = 1 'Add By Sindy 2015/9/17 結案中
   End If
'   If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Or Pub_StrUserSt03 = "M51" Then
'      LblCntTime.Visible = True
'   Else
'      LblCntTime.Visible = False
'   End If
   '2014/6/9 END
   
   'Added by Morgan 2016/3/1
   m_AttachPath = App.path & "\" & strUserNum
   KillTemp
   'end 2016/3/1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If pub_CallNextForm = True Then
'      strSql = "select * from executelog where el01='frm210125' and el02='" & strUserNum & "' and el03=" & strSrvDate(1)
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI <> 1 Then
         pub_CallNextForm = True
         frm210125.Show
         frm210125.cmdSearch_Click
'      End If
   End If
   Set frm100123 = Nothing
End Sub

'Add By Sindy 2014/6/19
Private Sub GrdDataList_Click()
Dim nCol As Integer, nRow As Integer
   With grdDataList
      .Visible = False
      nCol = .MouseCol
      nRow = .MouseRow
      If nRow = 0 And .TextMatrix(nRow, nCol) <> "V" Then
         .col = nCol
         If m_blnColOrderAsc = False Then '字串降冪
            .Sort = 5 '字串昇冪
            m_blnColOrderAsc = True
         Else
            .Sort = 6 '字串降冪
            m_blnColOrderAsc = False
         End If
      End If
      .Visible = True
   End With
End Sub

Private Sub grdDataList_SelChange()
Dim nCol As Integer, nRow As Integer

   grdDataList.Visible = False
   grdDataList.row = grdDataList.MouseRow
   grdDataList.col = 0
   If grdDataList.row <> 0 Then
      If grdDataList.Text = "V" Then
         grdDataList.Text = ""
         Call recovercolor(grdDataList.row) 'Modify By Sindy 2014/6/25 Sindy 還原顏色
'         For i = 0 To GrdDataList.Cols - 1
'            GrdDataList.col = i
'            If GrdDataList.CellBackColor = &HFFC0C0 Then
'               GrdDataList.CellBackColor = &H80000018
'            Else
'               GrdDataList.CellBackColor = &H8080FF
'            End If
'         Next i
      Else
         grdDataList.Text = "V"
         For i = 0 To grdDataList.Cols - 1
            grdDataList.col = i
            If grdDataList.CellBackColor = &H80000018 Then
               grdDataList.CellBackColor = &HFFC0C0
            Else
               grdDataList.CellBackColor = &HC0&
            End If
         Next i
      End If
   End If
   grdDataList.Visible = True
   
   SetEMailEnable 'Added by Morgan 2016/3/2
End Sub

'add by nickc 2007/01/18
Private Sub systemkind_GotFocus()
   TextInverse systemkind
   'edit by nickc 2007/06/06 切換輸入法改用API
   'systemkind.IMEMode = 2
   CloseIme
End Sub

Private Sub systemkind_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCloseDate_GotFocus(Index As Integer)
   'If Index = 1 Then txtCloseDate(Index) = txtCloseDate(Index - 1)
   TextInverse txtCloseDate(Index)
   'edit by nickc 2007/06/06 切換輸入法改用API
   'txtCloseDate(Index).IMEMode = 2
   CloseIme
End Sub

Private Sub txtCloseDate_LostFocus(Index As Integer)
   If txtCloseDate(Index).Tag <> txtCloseDate(Index).Text Then
      Combo2.ListIndex = 4 '3 設n個月 為空白
   End If
End Sub

Private Sub txtCloseDate_Validate(Index As Integer, Cancel As Boolean)
   If txtCloseDate(Index) <> "" Then
      If ChkDate(txtCloseDate(Index)) = False Then
         Cancel = True
         txtCloseDate(Index).SetFocus
         txtCloseDate_GotFocus Index
         'add by nickc 2005/08/15
         Exit Sub
      End If
      'add by nickc 2005/08/15
      If Index = 1 Then
         If RunNick2(txtCloseDate(0), txtCloseDate(1)) = True Then
            txtCloseDate(Index).SetFocus
            txtCloseDate_GotFocus Index
            Cancel = True
            Exit Sub
         End If
      End If
'      txtCloseDate(Index).Tag = txtCloseDate(Index).Text
   End If
End Sub

'add by nickc 2007/01/18
Private Sub txtCP10_GotFocus()
   TextInverse txtCP10
   'edit by nickc 2007/06/06 切換輸入法改用API
   'txtCP10.IMEMode = 2
   CloseIme
End Sub

'add by nickc 2007/01/18
Private Sub txtCU1_GotFocus()
   TextInverse txtCU1
   'edit by nickc 2007/06/06 切換輸入法改用API
   'txtCU1.IMEMode = 2
   CloseIme
End Sub

Private Sub txtCU1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCU1_LostFocus()
   If Trim(txtCU1) <> "" Then
      txtCU1 = Mid(txtCU1 & "000000000", 1, 9)
   End If
End Sub

Private Sub txtCU2_GotFocus()
   If Len(txtCU1) = 9 Then
      txtCU2 = Left(txtCU1, 6) & "ZZZ"
      txtCU2.SelStart = 6
      txtCU2.SelLength = 3
   End If
   TextInverse txtCU2
   'edit by nickc 2007/06/06 切換輸入法改用API
   'txtCU2.IMEMode = 2
   CloseIme
End Sub

Private Sub txtCU2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSales_Change()
   If Len(txtSales) > 4 Then
      lblSalesName = StaffQuery(Trim(txtSales))
   Else
      lblSalesName = ""
   End If
End Sub

Private Sub txtSales_GotFocus()
   TextInverse txtSales
   'edit by nickc 2007/06/06 切換輸入法改用API
   'txtSales.IMEMode = 2
   CloseIme
   If Combo3.Enabled = True And Combo3.Visible = True Then Combo3.SetFocus 'Add By Sindy 2016/5/4
End Sub

'Add By Sindy 2010/11/26
Private Sub txtSales_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSales_LostFocus()
   'add by nickc 2006/11/21
'   If Mid(txtSalesArea, 1, 1) <> "S" And Mid(txtSalesArea1, 1, 1) <> "S" Then
'       chkNP.Enabled = False
'       chkNP.Value = 0
'   Else
'       chkNP.Enabled = True
'   End If
   If Trim(txtSales) = "" Then
       lblSalesName = ""
   End If
End Sub

Private Sub txtSales_Validate(Cancel As Boolean)
   'Modify By Sindy 2023/9/6
   If PUB_txtSales_Limit(txtSales, m_strListPer, txtZone, txtSalesArea, txtSalesArea1, _
                         bolSpecMan, strSpecCode, lblSalesName) = False Then
      If txtSales.Visible = True Then 'Added by Lydia 2021/05/20 排除隱藏
        txtSales.SetFocus
        txtSales_GotFocus
      End If 'Added by Lydia 2021/05/20
      Cancel = True
      Exit Sub
   End If
   '2023/9/6 END
End Sub

Private Sub txtSalesArea_GotFocus()
   TextInverse txtSalesArea
   'edit by nickc 2007/06/06 切換輸入法改用API
   'txtSalesArea.IMEMode = 2
   CloseIme
End Sub

Private Sub txtSalesArea_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesArea_LostFocus()
   'add by nickc 2006/11/21
'   If Mid(txtSalesArea, 1, 1) <> "S" And Mid(txtSalesArea1, 1, 1) <> "S" Then
'       chkNP.Enabled = False
'       chkNP.Value = 0
'   Else
'       chkNP.Enabled = True
'   End If
End Sub

Private Sub txtSalesArea_Validate(Cancel As Boolean)
   'add by nickc 2006/11/21
'   If Mid(txtSalesArea, 1, 1) <> "S" And Mid(txtSalesArea1, 1, 1) <> "S" Then
'       chkNP.Enabled = False
'       chkNP.Value = 0
'   Else
'       chkNP.Enabled = True
'   End If
End Sub

Private Sub txtSalesArea1_GotFocus()
   TextInverse txtSalesArea1
   'edit by nickc 2007/06/06 切換輸入法改用API
   'txtSalesArea.IMEMode = 2
   CloseIme
End Sub

Private Sub txtSalesArea1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesArea1_LostFocus()
   'add by nickc 2006/11/21
'   If Mid(txtSalesArea, 1, 1) <> "S" And Mid(txtSalesArea1, 1, 1) <> "S" Then
'       chkNP.Enabled = False
'       chkNP.Value = 0
'   Else
'       chkNP.Enabled = True
'   End If
End Sub

'add by nickc 2005/08/15
Private Sub txtSalesArea1_Validate(Cancel As Boolean)
   If Trim(txtSalesArea1) <> "" Then
      If RunNick(txtSalesArea, txtSalesArea1) = True Then
         Cancel = True
         Exit Sub
      End If
   End If
   'add by nickc 2006/11/21
'   If Mid(txtSalesArea, 1, 1) <> "S" And Mid(txtSalesArea1, 1, 1) <> "S" Then
'       chkNP.Enabled = False
'       chkNP.Value = 0
'   Else
'       chkNP.Enabled = True
'   End If
End Sub

Private Sub txtZone_GotFocus()
   TextInverse txtZone
   'edit by nickc 2007/06/06 切換輸入法改用API
   'txtZone.IMEMode = 2
   CloseIme
End Sub

Private Sub txtZone_KeyPress(KeyAscii As Integer)
   If (KeyAscii < Asc("1") Or KeyAscii > Asc("4")) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Public Sub PubShowNextData()
Dim i As Integer, j As Integer
Dim strCust As String   '2011/8/9 add by sonia
'Add by Amy 2018/05/24
Dim intCount As Integer '計算筆數
Dim strChoose(9) As String  '勾選資料內容
Dim nFrm As Form 'Added by Lydia 2021/04/13
Dim arrKey() As String, strCaseNo(4) As String, strNextYearDesc As String 'Added by Morgan 2021/9/1
Dim oRunform As Form 'Add By Sindy 2022/9/16
   
   'Add By Sindy 2022/9/16
   If strSrvDate(1) >= 接洽單電子收文啟用日 Then
      Set oRunform = frm090801_New
   Else
      Set oRunform = frm090801
   End If
   '2022/9/16 END
   
'Memo by Lydia 2021/05/10 下列欄位值設定TextMatrix(i,x) 改為變數
'colPKey 'PKey，原位置21
'colType '事件，原位置5
'colCaseNo '本所案號，原位置9
'colCP06 '本所期限，原位置7
'colCPM '案件性質，原位置12
'colNP01 '收文號，原位置19
'colNP22 '序號，原位置20
'end 2021/05/10

SetEMailEnable 'Added by Morgan 2016/3/2

   Select Case cmdState
      Case 0 '案件基本資料
         Me.Enabled = False
         For i = 1 To grdDataList.Rows - 1
            grdDataList.col = 0
            grdDataList.row = i
            If Trim(grdDataList.Text) = "V" Then
               Dim Str01 As String
               grdDataList.col = 0
               grdDataList.Text = ""
               For j = 0 To grdDataList.Cols - 1
                  grdDataList.col = j
                  If grdDataList.CellBackColor = &HFFC0C0 Then
                    grdDataList.CellBackColor = &H80000018
                  Else
                    grdDataList.CellBackColor = &H8080FF
                  End If
               Next j
               
               'Add By Sindy 2010/01/15
               If Trim(grdDataList.TextMatrix(i, colType)) = "行事曆" Then '事件
                  If fnSaveParentForm(Me) = False Then
                     Me.Enabled = True
                     Exit Sub
                  End If
                  Screen.MousePointer = vbHourglass
                  frm100101_23.Show
                  frm100101_23.Tag = Pub_RplStr(Trim(grdDataList.TextMatrix(i, colPKey))) 'PKey
                  frm100101_23.CmdOk1(1).Visible = False 'Add By Sindy 2014/6/17 結束按鈕隱藏
                  frm100101_23.StrMenu
                  Screen.MousePointer = vbDefault
                  Me.Enabled = True
                  Exit Sub
               Else
               '2010/01/15 End
                  'Modify By Sindy 2012/8/24
                  'grdDataList.col = 8
                  grdDataList.col = colCaseNo '本所案號
                  '2012/8/24 End
                  Str01 = SystemNumber(grdDataList, 1)
                  If Mid(UCase(Str01), 1, 1) = "N" Then
                     Str01 = Mid(Str01, 2, 3)
                  End If
                  If Not IsNull(grdDataList.Text) Then
                     If fnSaveParentForm(Me) = False Then
                        Me.Enabled = True
                        Exit Sub
                     End If
                     Select Case Str01
                         Case "CFP", "FCP", "P"   '專利
                               Screen.MousePointer = vbHourglass
                               frm100101_3.Show
                               frm100101_3.Tag = Pub_RplStr(grdDataList.Text)
                               frm100101_3.cmdOK(1).Visible = False 'Add By Sindy 2014/6/17 結束按鈕隱藏
                               frm100101_3.StrMenu
                               Screen.MousePointer = vbDefault
                         Case "CFT", "FCT", "T", "TF"   '商標
                               Screen.MousePointer = vbHourglass
                               frm100101_4.Show
                               frm100101_4.Tag = Pub_RplStr(grdDataList.Text)
                               frm100101_4.cmdOK(3).Visible = False 'Add By Sindy 2014/6/17 結束按鈕隱藏
                               frm100101_4.StrMenu
                               Screen.MousePointer = vbDefault
                         'Modify By Sindy 2009/07/24 增加LIN系統類別
                         'modify by sonia 2019/7/30 +ACS系統類別
                         Case "CFL", "FCL", "L", "LIN", "ACS"     '法務
                               Screen.MousePointer = vbHourglass
                               frm100101_5.Show
                               frm100101_5.Tag = Pub_RplStr(grdDataList.Text)
                               frm100101_5.cmdOK(3).Visible = False 'Add By Sindy 2014/6/17 結束按鈕隱藏
                               frm100101_5.StrMenu
                               Screen.MousePointer = vbDefault
                         Case "LA"            '顧問
                               Screen.MousePointer = vbHourglass
                               frm100101_6.Show
                               frm100101_6.Tag = Pub_RplStr(grdDataList.Text)
                               frm100101_6.cmdOK(1).Visible = False 'Add By Sindy 2014/6/17 結束按鈕隱藏
                               frm100101_6.StrMenu
                               Screen.MousePointer = vbDefault
                         Case Else                  '服務
                              Select Case Str01
                                  Case "TB"    '條碼
                                     Screen.MousePointer = vbHourglass
                                     frm100101_7.Show
                                     frm100101_7.Tag = Pub_RplStr(grdDataList.Text)
                                     frm100101_7.cmdOK(2).Visible = False 'Add By Sindy 2014/6/17 結束按鈕隱藏
                                     frm100101_7.StrMenu
                                     Screen.MousePointer = vbDefault
                                  Case "TM"
                                     Screen.MousePointer = vbHourglass
                                     frm100101_8.Show
                                     frm100101_8.Tag = Pub_RplStr(grdDataList.Text)
                                     frm100101_8.cmdOK(3).Visible = False 'Add By Sindy 2014/6/17 結束按鈕隱藏
                                     frm100101_8.StrMenu
                                     Screen.MousePointer = vbDefault
                                  Case "TD"
                                     Screen.MousePointer = vbHourglass
                                     frm100101_9.Show
                                     frm100101_9.Tag = Pub_RplStr(grdDataList.Text)
                                     frm100101_9.cmdOK(2).Visible = False 'Add By Sindy 2014/6/17 結束按鈕隱藏
                                     frm100101_9.StrMenu
                                     Screen.MousePointer = vbDefault
                                  Case "TC", "CFC"
                                     Screen.MousePointer = vbHourglass
                                     frm100101_A.Show
                                     frm100101_A.Tag = Pub_RplStr(grdDataList.Text)
                                     frm100101_A.cmdOK(2).Visible = False 'Add By Sindy 2014/6/17 結束按鈕隱藏
                                     frm100101_A.StrMenu
                                     Screen.MousePointer = vbDefault
                                  Case Else
                                     Screen.MousePointer = vbHourglass
                                     frm100101_B.Show
                                     frm100101_B.Tag = Pub_RplStr(grdDataList.Text)
                                     frm100101_B.cmdOK(2).Visible = False 'Add By Sindy 2014/6/17 結束按鈕隱藏
                                     frm100101_B.StrMenu
                                     Screen.MousePointer = vbDefault
                               End Select
                     End Select
                  End If
               End If
               Call recovercolor(i) '2011/8/9 add by sonia 還原顏色
               Me.Enabled = True
               Exit Sub
            End If
         Next i
         Me.Enabled = True
      Case 1 '案件進度
         Me.Enabled = False
         StrTag = ""
         For i = 1 To grdDataList.Rows - 1
            grdDataList.col = 0
            grdDataList.row = i
            If Trim(grdDataList.Text) = "V" Then
               grdDataList.col = 0
               grdDataList.Text = ""
               For j = 0 To grdDataList.Cols - 1
                  grdDataList.col = j
                  If grdDataList.CellBackColor = &HFFC0C0 Then
                     grdDataList.CellBackColor = &H80000018
                  Else
                    grdDataList.CellBackColor = &H8080FF
                  End If
               Next j
               
               'Add By Sindy 2010/01/15
               If Trim(grdDataList.TextMatrix(i, colType)) = "行事曆" Then '事件
                  If fnSaveParentForm(Me) = False Then
                     Me.Enabled = True
                     Exit Sub
                  End If
                  Screen.MousePointer = vbHourglass
                  frm100101_23.Show
                  frm100101_23.Tag = Pub_RplStr(Trim(grdDataList.TextMatrix(i, colPKey))) 'PKey
                  frm100101_23.CmdOk1(1).Visible = False 'Add By Sindy 2014/6/17 結束按鈕隱藏
                  frm100101_23.StrMenu
                  Screen.MousePointer = vbDefault
'                  Me.Enabled = True
'                  Exit Sub
               Else
               '2010/01/15 End
                  'Modify By Sindy 2012/8/24
                  'grdDataList.col = 8
                  grdDataList.col = colCaseNo '本所案號
                  '2012/8/24 End
                  If Not IsNull(grdDataList.Text) Then
                     If fnSaveParentForm(Me) = False Then
                        Me.Enabled = True
                        Exit Sub
                     End If
                     Screen.MousePointer = vbHourglass
                     frm100101_2.Show
                     frm100101_2.Tag = Pub_RplStr(grdDataList.Text)
                     frm100101_2.cmdOK(6).Visible = False 'Add By Sindy 2014/6/17 結束按鈕隱藏
                     frm100101_2.StrMenu
                     Screen.MousePointer = vbDefault
                  End If
               End If
               Call recovercolor(i) '2011/8/9 add by sonia 還原顏色
               Me.Enabled = True
               Exit Sub
            End If
         Next i
         Me.Enabled = True
         
      'add by nickc 2005/09/08
      Case 2 '內商延展結案通知
         Me.Enabled = False
         Screen.MousePointer = vbHourglass
         grdDataList.MousePointer = flexArrowHourGlass
         'add by nickc 2005/09/20 計算筆數
         StrTmpCp01020304 = ""
         StrCompCp01020304 = ""
         intCount = 0 'Add by Amy 2018/05/24
         For i = 1 To grdDataList.Rows - 1
            grdDataList.col = 0
            grdDataList.row = i
            If Trim(grdDataList.Text) = "V" Then
               'Modify By Sindy 2012/8/24
               'grdDataList.col = 8
               grdDataList.col = colCaseNo '本所案號
               '2012/8/24 End
              'edit by nickc 2006/05/12
              'If SystemNumber(grdDataList.Text, 1) = "T" And grdDataList.TextMatrix(i, coltype) = "未處理" And Mid(grdDataList.TextMatrix(i, 6), 1, 1) <> "*" Then
              'Modify By Sindy 2010/4/19
      '        If SystemNumber(grdDataList.Text, 1) = "T" And grdDataList.TextMatrix(i, coltype) = "未處理" Then
               If (SystemNumber(grdDataList.Text, 1) = "T" Or SystemNumber(grdDataList.Text, 1) = "TF") And _
                  (grdDataList.TextMatrix(i, colType) = "未處理" Or grdDataList.TextMatrix(i, colType) = "未函知") Then '事件
                  If Mid(grdDataList.TextMatrix(i, colCP06), 1, 1) <> "*" Then '本所期限
                      'Modify By Sindy 2012/8/24
                      'grdDataList.col = 11
                      grdDataList.col = colCPM '案件性質
                      '2012/8/24 End
                      'edit by nickc 2006/06/12加入第二期
                      'If Trim(grdDataList.Text) = "延展" Or Trim(grdDataList.Text) = "續展" Then
                      If Trim(grdDataList.Text) = "延展" Or Trim(grdDataList.Text) = "續展" Or Trim(grdDataList.Text) = "第二期註冊費" Then
                             StrTmpCp01020304 = StrTmpCp01020304 & grdDataList.TextMatrix(i, colCaseNo) & vbCrLf '本所案號
                             StrCompCp01020304 = StrCompCp01020304 & grdDataList.TextMatrix(i, colCaseNo) & ","
                      'add by nickc 2005/09/29
                      Else
                          'edit by nickc 2006/06/12
                          'MsgBox "含有非內商延展案，請重新點選！", vbExclamation, "發生錯誤！"
                          MsgBox "含有非內商延展案或非內商第二期註冊費，請重新點選！", vbExclamation, "發生錯誤！"
                          grdDataList.MousePointer = flexDefault
                          Screen.MousePointer = vbDefault
                          Me.Enabled = True
                          Exit Sub
                      End If
                      'Add by Amy 2018/05/24 增加結案說明畫面
                      '判斷只能選10筆(從0開始算)
                      If intCount > 9 Then
                          MsgBox "最多只能點選10筆，請重新點選！", vbExclamation, "發生錯誤！"
                          grdDataList.MousePointer = flexDefault
                          Screen.MousePointer = vbDefault
                          Me.Enabled = True
                          Exit Sub
                      End If
                      strChoose(intCount) = GetValue(i, "本所案號") & "," & Trim(GetValue(i, "本所期限")) & "," & GetValue(i, "法定期限") & "," & _
                                                            GetValue(i, "案件名稱") & "," & GetValue(i, "收文號") & "," & GetValue(i, "序號")
                       intCount = intCount + 1
                      'end 2018/05/24
                  'add by nickc 2006/05/12
                  Else
                      'edit by nickc 2006/06/12
                      'MsgBox "含有已通知結案延展資料，請重新點選！", vbExclamation, "發生錯誤！"
                      MsgBox "含有已通知結案延展或第二期註冊費資料，請重新點選！", vbExclamation, "發生錯誤！"
                      grdDataList.MousePointer = flexDefault
                      Screen.MousePointer = vbDefault
                      Me.Enabled = True
                      Exit Sub
                  End If
               'add by nickc 2005/09/29
               Else
                   'edit by nickc 2006/06/12
                   'MsgBox "含有非內商延展案，請重新點選！", vbExclamation, "發生錯誤！"
                   MsgBox "含有非內商延展案非內商第二期註冊費，請重新點選！", vbExclamation, "發生錯誤！"
                   grdDataList.MousePointer = flexDefault
                   Screen.MousePointer = vbDefault
                   Me.Enabled = True
                   Exit Sub
               End If
            End If
         Next i
         If StrTmpCp01020304 <> "" Then
            'Modify by Amy 2018/05/24 增加商標延展結案說明輸入畫面
'            If MsgBox(StrTmpCp01020304 & vbCrLf & vbCrLf & "確定結案？", vbYesNo, "警告！") = vbNo Then
'               grdDataList.MousePointer = flexDefault
'               Screen.MousePointer = vbDefault
'               Me.Enabled = True
'               Exit Sub
'            End If
'            Call SaveT102inform 'Modify By Sindy 2015/3/10 改寫到Func裡
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            frm100123_1.SetParent Me
            frm100123_1.Show
            Call frm100123_1.StrMenu(strChoose(), intCount - 1)
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Exit Sub
         End If
         'end 2018/05/24
         grdDataList.MousePointer = flexDefault
         Screen.MousePointer = vbDefault
         Me.Enabled = True
      '2011/8/9 ADD BY SONIA
      'Remove by Lydia 2021/05/10 原本「最後收文人員」已不在使用，改為「管制備註作業」
'      Case 3 '最後收文人員
'         Me.Enabled = False
'         Screen.MousePointer = vbHourglass
'         grdDataList.MousePointer = flexArrowHourGlass
'         For i = 1 To grdDataList.Rows - 1
'            grdDataList.col = 0
'            grdDataList.row = i
'            If Trim(grdDataList.Text) = "V" Then
'               grdDataList.col = 0
'               grdDataList.Text = ""
'               'Modify By Sindy 2012/8/24
'               'grdDataList.col = 15
'               grdDataList.col = 16 '申請人
'               '2012/8/24 End
'               strCust = grdDataList.Text
'               'Modify By Sindy 2012/8/24
'               'grdDataList.col = 8
'               grdDataList.col = colCaseNo '本所案號
'               '2012/8/24 End
'               '抓該客戶所有案件最後收文之智權人員,包含離職人員
'               If grdDataList.Text <> "" Then
'                  strExc(0) = "select st02 from staff,(select max(cp05||cp09||cp13) cp13 from ( " & _
'                              "      Select cp05,cp09,cp13 From patent, caseprogress Where pa26='" & GetPrjPeopleNum1(grdDataList.Text) & "' and pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04 and cp09<'B' " & _
'                              "union Select cp05,cp09,cp13 From trademark, caseprogress Where tm23='" & GetPrjPeopleNum1(grdDataList.Text) & "' and tm01=cp01 and tm02=cp02 and tm03=cp03 and tm04=cp04 and cp09<'B' " & _
'                              "union Select cp05,cp09,cp13 From lawcase, caseprogress Where lc11='" & GetPrjPeopleNum1(grdDataList.Text) & "' and lc01=cp01 and lc02=cp02 and lc03=cp03 and lc04=cp04 and cp09<'B' " & _
'                              "union Select cp05,cp09,cp13 From servicepractice, caseprogress Where sp08='" & GetPrjPeopleNum1(grdDataList.Text) & "' and sp01=cp01 and sp02=cp02 and sp03=cp03 and sp04=cp04 and cp09<'B' " & _
'                              "union Select cp05,cp09,cp13 From hirecase, caseprogress Where hc05='" & GetPrjPeopleNum1(grdDataList.Text) & "' and hc01=cp01 and hc02=cp02 and hc03=cp03 and hc04=cp04 and cp09<'B' " & _
'                              ")) aa where substr(aa.cp13,18)=st01(+) "
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI = 1 Then
'                      MsgBox strCust & "(" & grdDataList.Text & ")" & "所有案件最後收文智權人員為 " & RsTemp.Fields(0).Value & " ！"
'                  End If
'               End If
'               '還原顏色
'               Call recovercolor(i)
'            End If
'         Next i
'         grdDataList.MousePointer = flexDefault
'         Screen.MousePointer = vbDefault
'         Me.Enabled = True
'      '2011/8/9 END

      'Added by Lydia 2021/05/10
      Case 3  '管制備註
          Set nFrm = Forms(0).GetForm("frm100123_2")
          If nFrm Is Nothing Then
              Exit Sub
          End If
          Me.Enabled = False
          StrTag = ""
          strExc(1) = "": strExc(2) = "": strExc(3) = "": strExc(4) = ""
          For i = 1 To grdDataList.Rows - 1
             grdDataList.col = 0
             grdDataList.row = i
             'Memo by Lydia 2021/05/28 保留單筆的寫法
'             If Trim(grdDataList.Text) = "V" Then
'                 grdDataList.col = 0
'                 grdDataList.Text = ""
'                 For j = 0 To grdDataList.Cols - 1
'                    grdDataList.col = j
'                    If grdDataList.CellBackColor = &HFFC0C0 Then
'                       grdDataList.CellBackColor = &H80000018
'                    Else
'                      grdDataList.CellBackColor = &H8080FF
'                    End If
'                 Next j
'                 If "" & grdDataList.TextMatrix(i, colNP01) <> "" And Val("" & grdDataList.TextMatrix(i, colNP22)) > 0 Then
'                     Screen.MousePointer = vbHourglass
'                     Call nFrm.SetParent(Me, Trim("" & grdDataList.TextMatrix(i, colNP01)), Trim("" & grdDataList.TextMatrix(i, colNP22)))
'                     nFrm.Show
'                     Me.Hide
'                     Screen.MousePointer = vbDefault
'                 End If
'                 Call recovercolor(i) '還原顏色
'                 Me.Enabled = True
'                 Exit Sub
'             End If
             'end 2021/05/28
             '可選多筆連續輸入管制備註，存檔後自動帶下一筆；取消則回前畫面，但有勾選但未顯示的資料的勾選符號必須保留，這樣才知道做到哪一筆。
              '--------讀取全部資料,分別記錄勾選列數、收文號、序號
              If Trim(grdDataList.Text) = "V" Then
                strExc(1) = strExc(1) & "," & Format(i, "000")
                strExc(2) = strExc(2) & "," & Trim("" & grdDataList.TextMatrix(i, colNP01))
                strExc(3) = strExc(3) & "," & Trim("" & grdDataList.TextMatrix(i, colNP22))
              End If
         Next i
         
         If strExc(1) <> "" Then
             strExc(1) = Mid(strExc(1), 2)
             strExc(2) = Mid(strExc(2), 2)
             strExc(3) = Mid(strExc(3), 2)
             Call nFrm.SetParent(Me, strExc(1), strExc(2), strExc(3))
             nFrm.Show
             Me.Hide
             Screen.MousePointer = vbDefault
         End If
         Me.Enabled = True
      'end 2021/05/10
      'Add By Sindy 2014/6/3
      Case 4 '卷宗區
         Me.Enabled = False
         StrTag = ""
         For i = 1 To grdDataList.Rows - 1
            grdDataList.col = 0
            grdDataList.row = i
            If Trim(grdDataList.Text) = "V" Then
               grdDataList.col = 0
               grdDataList.Text = ""
               For j = 0 To grdDataList.Cols - 1
                  grdDataList.col = j
                  If grdDataList.CellBackColor = &HFFC0C0 Then
                     grdDataList.CellBackColor = &H80000018
                  Else
                    grdDataList.CellBackColor = &H8080FF
                  End If
               Next j
               grdDataList.col = colCaseNo '本所案號
               If Not IsNull(grdDataList.Text) Then
                  'Modified by Morgan 2021/7/23 改先不隱藏，否則會觸發已開最上層表單的 Form_Active 事件
                  'If fnSaveParentForm(Me) = False Then
                  If fnSaveParentForm(Me, True) = False Then
                     Me.Enabled = True
                     Exit Sub
                  End If
                  Screen.MousePointer = vbHourglass
                  StrTag = Pub_RplStr(grdDataList.Text)
                  If UBound(Split(StrTag, "-")) = 1 Then
                     StrTag = StrTag & "-0-00"
                  End If
                  frm100101_L.m_strKey = StrTag
                  'frm100101_L.Hide
                  frm100101_L.SetParent Me
                  If frm100101_L.QueryData = True Then
                     frm100101_L.Show
                     'Me.Hide
                     Me.Hide 'Added by Morgan 2021/7/23
                  Else
                     Unload frm100101_L
                  End If
                  Screen.MousePointer = vbDefault
               End If
               Call recovercolor(i) '還原顏色
               Me.Enabled = True
               Exit Sub
            End If
         Next i
         Me.Enabled = True
      Case 5 '收文
         Me.Enabled = False
         StrTag = ""
         For i = 1 To grdDataList.Rows - 1
            grdDataList.col = 0
            grdDataList.row = i
            If Trim(grdDataList.Text) = "V" Then
               grdDataList.col = 0
               grdDataList.Text = ""
               For j = 0 To grdDataList.Cols - 1
                  grdDataList.col = j
                  If grdDataList.CellBackColor = &HFFC0C0 Then
                     grdDataList.CellBackColor = &H80000018
                  Else
                    grdDataList.CellBackColor = &H8080FF
                  End If
               Next j
               grdDataList.col = colCaseNo '本所案號
               If Not IsNull(grdDataList.Text) Then
                  If fnSaveParentForm(Me) = False Then
                     Me.Enabled = True
                     Exit Sub
                  End If
                  Screen.MousePointer = vbHourglass
                  StrTag = Pub_RplStr(grdDataList.Text)
                  If UBound(Split(StrTag, "-")) = 1 Then
                     StrTag = StrTag & "-0-00"
                  End If
                  'Modify By Sindy 2022/9/16 frm090801 改用 oRunform
                  'Added by Lydia 2021/04/13 查名單輸入的設定; (From 林青祺) 國內接洽單收文T台灣申請案，沒有顯示「查名單輸入」按鈕。
                  Set nFrm = Forms(0).GetForm("frm090126")
                  If Not nFrm Is Nothing Then
                     Set oRunform.Tmpfrm090126 = nFrm
                  End If
                  'end 2021/04/13
                  oRunform.bolExternalCall = True '記錄是外部程式呼叫使用
                  oRunform.SetParent Me
                  oRunform.Show
                  oRunform.Tag = StrTag
                  oRunform.Option1(1).Value = True
                  oRunform.Text1(6) = SystemNumber(StrTag, 1)
                  oRunform.Text1(7) = SystemNumber(StrTag, 2)
                  oRunform.Text1(8) = SystemNumber(StrTag, 3)
                  oRunform.Text1(9) = SystemNumber(StrTag, 4)
                  Call oRunform.Text1_LostFocus(9)
                  oRunform.bolExternalCall = False '還原預設值
                  Screen.MousePointer = vbDefault
               End If
               Call recovercolor(i) '還原顏色
               Me.Enabled = True
               Exit Sub
            End If
         Next i
         Me.Enabled = True
      '2014/6/3 END
      'Add By Sindy 2014/6/19
      Case 6 '結案(非延展案)
         Me.Enabled = False
         'Add By Sindy 2014/7/4
         For i = 1 To grdDataList.Rows - 1
            grdDataList.col = 0
            grdDataList.row = i
            If Trim(grdDataList.Text) = "V" Then
               grdDataList.col = colCaseNo '本所案號
               'Modify by Amy 2020/05/20 +結案中
               If (SystemNumber(grdDataList.Text, 1) = "T" Or SystemNumber(grdDataList.Text, 1) = "TF") And _
                  (grdDataList.TextMatrix(i, colType) = "結案中" Or grdDataList.TextMatrix(i, colType) = "未處理" Or grdDataList.TextMatrix(i, colType) = "未函知") Then '事件
                  grdDataList.col = colCPM '案件性質
                  'Modify by Amy 2020/05/18 +ChkT102Inform 內商延展、續展從此按鈕進入,若已結案彈訊息
                  If (Trim(grdDataList.Text) = "延展" Or Trim(grdDataList.Text) = "續展" Or Trim(grdDataList.Text) = "第二期註冊費") _
                   And ChkT102Inform(grdDataList.TextMatrix(i, colNP01), grdDataList.TextMatrix(i, colNP22)) = True Then
                     'MsgBox "含有內商延展案或內商第二期註冊費，請重新點選！", vbExclamation, "發生錯誤！"
                     MsgBox "此結案單已存在，不可重覆作業！", vbCritical, "操作錯誤！"
                     grdDataList.MousePointer = flexDefault
                     Screen.MousePointer = vbDefault
                     Me.Enabled = True
                     Exit Sub
                  End If
                  'end 2020/05/18
               End If
            End If
         Next i
         '2014/7/4 END
         StrTag = ""
         For i = 1 To grdDataList.Rows - 1
            grdDataList.col = 0
            grdDataList.row = i
            If Trim(grdDataList.Text) = "V" Then
               grdDataList.col = 0
               grdDataList.Text = ""
               For j = 0 To grdDataList.Cols - 1
                  grdDataList.col = j
                  If grdDataList.CellBackColor = &HFFC0C0 Then
                     grdDataList.CellBackColor = &H80000018
                  Else
                    grdDataList.CellBackColor = &H8080FF
                  End If
               Next j
               grdDataList.col = colCaseNo '本所案號
               If Not IsNull(grdDataList.Text) Then
                  If fnSaveParentForm(Me) = False Then
                     Me.Enabled = True
                     Exit Sub
                  End If
                  Screen.MousePointer = vbHourglass
                  StrTag = Pub_RplStr(grdDataList.Text)
                  If UBound(Split(StrTag, "-")) = 1 Then
                     StrTag = StrTag & "-0-00"
                  End If
                  frm210133.txt1(0) = SystemNumber(StrTag, 1)
                  frm210133.txt1(1) = SystemNumber(StrTag, 2)
                  frm210133.txt1(2) = SystemNumber(StrTag, 3)
                  frm210133.txt1(3) = SystemNumber(StrTag, 4)
                  frm210133.m_NP01 = grdDataList.TextMatrix(i, colNP01) 'Add By Sindy 2015/1/16
                  frm210133.m_NP22 = grdDataList.TextMatrix(i, colNP22) 'Add By Sindy 2015/1/16
                  frm210133.SetParent Me
                  If frm210133.doQuery = True Then
                     frm210133.Show
                  Else
                     Unload frm210133
                  End If
                  Screen.MousePointer = vbDefault
               End If
               Call recovercolor(i) '還原顏色
               Me.Enabled = True
               Exit Sub
            End If
         Next i
         Me.Enabled = True
      'Added by Morgan 2016/3/1
      Case 7 'EMail
         Me.Enabled = False
         cmdOK(7).Enabled = False
         StrTag = ""
         For i = 1 To grdDataList.Rows - 1
            grdDataList.col = 0
            grdDataList.row = i
            If Trim(grdDataList.Text) = "V" Then
               
               strExc(1) = GetValue(i, "收文號")
               strExc(2) = GetValue(i, "序號")
               strExc(3) = GetValue(i, "本所案號")
               strExc(4) = GetValue(i, "案件性質")
               strExc(5) = ""
               If Val(strExc(2)) > 0 Then
                  strExc(5) = GetMailAtt(strExc(1))
               End If
               
               If grdDataList.TextMatrix(i, colType) = "未處理" Then
                  strExc(6) = Trim(GetValue(i, "本所期限"))
                  strExc(7) = Trim(GetValue(i, "法定期限"))
                  
                  'Added by Morgan 2021/9/1 專利年費則於案件性質後加待繳費年度
                  If strExc(3) <> "" Then
                     arrKey = Split(strExc(3), "-")
                     strCaseNo(3) = "0": strCaseNo(4) = "00"
                     For intI = 0 To 3
                        If intI <= UBound(arrKey) Then
                           strCaseNo(intI + 1) = arrKey(intI)
                        End If
                     Next
                     If (strCaseNo(1) = "P" Or strCaseNo(1) = "CFP") And (strExc(4) = "年費" Or strExc(4) = "維持費" Or strExc(4) = "延展費") Then
                        PUB_GetNextYear strCaseNo, strNextYearDesc
                        strExc(4) = strExc(4) & IIf(strNextYearDesc = "", "", " [ " & strNextYearDesc & " ] ")
                     End If
                  End If
                  'end 2021/9/1
                  
                  PUB_ShowMailForm strExc(1), strExc(5), strExc(4), , strExc(3) & " " & strExc(4) & "期限提醒", strExc(6), strExc(7)
               Else
                  PUB_ShowMailForm strExc(1), strExc(5), strExc(4), , strExc(3) & " " & strExc(4) & "事宜"
               End If
               grdDataList.col = 0
               grdDataList.Text = ""
               For j = 0 To grdDataList.Cols - 1
                  grdDataList.col = j
                  If grdDataList.CellBackColor = &HFFC0C0 Then
                     grdDataList.CellBackColor = &H80000018
                  Else
                    grdDataList.CellBackColor = &H8080FF
                  End If
               Next j

               Call recovercolor(i) '還原顏色
               Me.Enabled = True
               Exit Sub
            End If
         Next i
         Me.Enabled = True
      'end 2016/3/1
      'Add by Amy 2020/03/25 資料寄信箱
      Case 8
        If SaveExcel = True Then
            PUB_SendMail strUserNum, strUserNum, "", Me.Caption & "-" & "資料寄送", "如摘要", , strExcelPath & xlsFileName
            MsgBox "寄信完成"
        End If
      '2014/6/19 END
      Case Else
   End Select
End Sub
'Added by Morgan 2016/3/1
Private Function GetMailAtt(pCP09 As String) As String
   Dim stFiles As String
   Dim stFileName As String
   Dim arrFileName() As String
   Dim idx As Integer
   Dim stFileNameList As String
   
   strExc(0) = "select cp09,cp10,cpm03,1 Srt from caseprogress,letterprogress,casepropertymap where cp43='" & pCP09 & "' and cp10='1913' and lp01(+)=cp09 and lp10='Y' and cpm01(+)=cp01 and cpm02(+)=cp10"
   strExc(0) = strExc(0) & " union all select cp09,cp10,cpm03,2 Srt from caseprogress,letterprogress,casepropertymap where cp09='" & pCP09 & "' and cp09>'C' and lp01(+)=cp09 and lp10='Y' and cpm01(+)=cp01 and cpm02(+)=cp10 order by Srt"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If PUB_GetAttachFile4Cust(RsTemp("cp09"), stFiles, m_AttachPath, False, RsTemp("cp10")) = True Then
         arrFileName = Split(stFiles, ";")
         For idx = UBound(arrFileName) To LBound(arrFileName) Step -1
            If arrFileName(idx) <> "" Then
               stFileName = m_AttachPath & "\" & arrFileName(idx)
               stFileNameList = stFileName & ";" & stFileNameList
            End If
         Next
      End If
   End If
   GetMailAtt = stFileNameList
End Function

'Added by Morgan 2016/3/1
Private Sub SetEMailEnable()
   Dim iRow As Integer, iCount As Integer
   
   cmdOK(7).Enabled = False
   iCount = 0
   With grdDataList
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) = "V" Then
         'modify by sonia 2019/6/11 +本所期限為2個月(專利年費2個月前才函知)內的未函知,T-220950移轉2019/4接進來案件,延展期限2019/8不會再函知
         'Modified by Lydia 2021/05/10 改用變數
         'If grdDataList.TextMatrix(iRow, 5) = "未處理" Or (grdDataList.TextMatrix(iRow, 5) = "未函知" And Val(ChangeTDateStringToTString(grdDataList.TextMatrix(iRow, 7))) <= Val(ChangeWDateStringToTString(DateAdd("m", 2, ChangeWStringToWDateString(strSrvDate(1)))))) Then
         If grdDataList.TextMatrix(iRow, colType) = "未處理" _
            Or (grdDataList.TextMatrix(iRow, colType) = "未函知" And Val(ChangeTDateStringToTString(Trim(grdDataList.TextMatrix(iRow, colCP06)))) <= Val(ChangeWDateStringToTString(DateAdd("m", 2, ChangeWStringToWDateString(strSrvDate(1)))))) Then
            cmdOK(7).Enabled = True
         Else
            cmdOK(7).Enabled = False
            Exit For
         End If
         iCount = iCount + 1
         '勾選兩筆以上不可EMail
         If iCount > 1 Then
            cmdOK(7).Enabled = False
            Exit For
         End If
      End If
   Next
   End With
End Sub

'Mark by Amy 2018/05/24 改至frm100123_1做
'Modify By Sindy 2015/3/10 改寫到Func裡,增加可以存放回覆單的功能
'Private Sub SaveT102inform()
'Dim strCaseNo As String
'Dim strCP01  As String, strCP02  As String, strCP03 As String, strCP04 As String
'Dim strFileName As String, intRow As Integer, bolSaveFile As Boolean
'Dim intHaveRf As Integer, strRFile As String, strChkR_Cp01020304 As String
'
'   '清除GrdAtt資料
'   For i = 0 To GrdAtt.Cols - 1
'      GrdAtt.TextMatrix(1, i) = ""
'   Next i
'   For i = GrdAtt.Rows - 1 To 2 Step -1
'      Call GrdAtt.RemoveItem(i)
'   Next i
'
'   bolSaveFile = False
'   TmpArrCaseNo = Split(StrCompCp01020304, ",")
'   If MsgBox("是否匯入回覆單？" & IIf(UBound(TmpArrCaseNo) > 1, vbCrLf & vbCrLf & "（註：下一視窗中，若同時按住＜Ctrl＞鍵，可點選多個檔案）", ""), vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
'AgainOpenAtt:
'      If OpenAddAtt = False Then
'         Exit Sub
'      Else
'         bolSaveFile = True
'      End If
'   End If
'
''on error GoTo CheckingErr
'
'AgainSaveAtt:
'   For i = 1 To grdDataList.Rows - 1
'      grdDataList.col = 0
'      grdDataList.row = i
'      If Trim(grdDataList.Text) = "V" Then
'         grdDataList.col = 9 '本所案號
'         If (SystemNumber(grdDataList.Text, 1) = "T" Or SystemNumber(grdDataList.Text, 1) = "TF") And _
'            (grdDataList.TextMatrix(i, 5) = "未處理" Or grdDataList.TextMatrix(i, 5) = "未函知") And Mid(grdDataList.TextMatrix(i, 7), 1, 1) <> "*" Then
'            grdDataList.col = 12 '案件性質
'            If Trim(grdDataList.Text) = "延展" Or Trim(grdDataList.Text) = "續展" Or Trim(grdDataList.Text) = "第二期註冊費" Then
'               StrTmpCp09 = grdDataList.TextMatrix(i, 19)
'               StrTmpNp22 = grdDataList.TextMatrix(i, 20)
'               strCaseNo = grdDataList.TextMatrix(i, 9)
'               strCP01 = SystemNumber(strCaseNo, 1)
'               strCP02 = SystemNumber(strCaseNo, 2)
'               strCP03 = SystemNumber(strCaseNo, 3)
'               strCP04 = SystemNumber(strCaseNo, 4)
'
'               '更新資料
'               CheckOC3
'               AdoRecordSet3.CursorLocation = adUseClient
'               AdoRecordSet3.Open "select * from t102inform where ti01=to_number(to_char(sysdate, 'YYYYMMDD')) and ti02='" & StrTmpCp09 & "' and ti04=" & StrTmpNp22, cnnConnection, adOpenStatic, adLockReadOnly
'               If AdoRecordSet3.RecordCount = 0 Then
'                  cnnConnection.BeginTrans
'
'                  '檢查是否有回覆單
'                  strFileName = ""
'                  intRow = 0
'                  If bolSaveFile = True Then
'                     '檢查是否有符合此案號的回覆單
'                     For intRow = 1 To GrdAtt.Rows - 1
'                        If GrdAtt.TextMatrix(intRow, 2) = strCaseNo And _
'                           (GrdAtt.TextMatrix(intRow, 3) <> "Y" And GrdAtt.TextMatrix(intRow, 3) <> "R") Then
'                           strFileName = GrdAtt.TextMatrix(intRow, 0)
'                           Exit For
'                        End If
'                     Next intRow
'                  End If
'                  If strFileName <> "" Then
'                     '檢查是否已有此案號回覆單
'                     'Modify By Sindy 2015/5/18
'                     'If PUB_ChkIsReplyFile(strCP01, strCP02, strCP03, strCP04) = True Then
'                     If PUB_ChkIsReplyFile(strCP01 & strCP02 & strCP03 & strCP04, , , , StrTmpNp22) = True Then
'                     '2015/5/18 END
'                        GrdAtt.TextMatrix(intRow, 3) = "R" '重覆
'                        GoTo ReadNext
'                     End If
'                     '存回覆單
'                     'Modify By Sindy 2015/5/18
'                     'If PUB_UpdReplyFile(strFileName, "", strCP01, strCP02, strCP03, strCP04) = False Then
'                     If PUB_UpdReplyFile(strFileName, "", strCP01, strCP02, strCP03, strCP04, , StrTmpNp22) = False Then
'                     '2015/5/18 END
'                        Exit Sub
'                     Else
'                        GrdAtt.TextMatrix(intRow, 3) = "Y" '已存
'                     End If
'                  End If
'
'                  '存結案記錄 : 1.沒有要存回覆單 或是 2.要存回覆單且有對應到電子檔
'                  If Not (bolSaveFile = True And strFileName = "") Then
'                     cnnConnection.Execute "insert into t102inform (ti01,ti02,ti03,ti04) values (to_number(to_char(sysdate, 'YYYYMMDD')),'" & StrTmpCp09 & "','" & strUserNum & "'," & StrTmpNp22 & ") "
'
'                     '已結案清除案號記錄
'                     StrCompCp01020304 = Replace(StrCompCp01020304, strCaseNo & ",", "")
'                     Call PUB_DelPCOrgFile(strFileName) '刪除原檔
'
'                     grdDataList.col = 7 '本所期限
'                     grdDataList.Text = "*" & grdDataList.TextMatrix(i, 7)
'                     grdDataList.col = 0
'                     grdDataList.Text = ""
'                     For j = 0 To grdDataList.Cols - 1
'                        grdDataList.col = j
'                        If grdDataList.CellBackColor = &HFFC0C0 Then
'                           grdDataList.CellBackColor = &H80000018 '反白
'                        Else
'                           grdDataList.CellBackColor = &H8080FF '變紅
'                        End If
'                     Next j
'                  End If
'                  cnnConnection.CommitTrans
'               End If
'            End If
'         End If
'      End If
'ReadNext:
'   Next i
'
'   '檢查有電子檔的檔案狀況
'   intRow = 0
'   intHaveRf = 0
'   strRFile = ""
'   strChkR_Cp01020304 = StrCompCp01020304
'   If GrdAtt.TextMatrix(1, 0) <> "" Then
'      For intRow = 1 To GrdAtt.Rows - 1
'         If GrdAtt.TextMatrix(intRow, 3) = "R" Then
'            intHaveRf = intHaveRf + 1
'            strRFile = strRFile & GrdAtt.TextMatrix(intRow, 1) & vbCrLf
'            strChkR_Cp01020304 = Replace(strChkR_Cp01020304, GrdAtt.TextMatrix(intRow, 2) & ",", "")
'         End If
'      Next intRow
'   End If
'   If strChkR_Cp01020304 <> "" Then
'      TmpArrCaseNo = Split(strChkR_Cp01020304, ",")
'      If MsgBox(Left(strChkR_Cp01020304, Len(strChkR_Cp01020304) - 1) & vbCrLf & vbCrLf & "無對應的電子檔，確定是否要重新點選回覆單？" & IIf(UBound(TmpArrCaseNo) > 1, vbCrLf & vbCrLf & "（註：下一視窗中，若同時按住＜Ctrl＞鍵，可點選多個檔案）", ""), vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
'         '重新點選回覆單
'         GoTo AgainOpenAtt
'      Else
'         bolSaveFile = False '不存回覆單
'         GoTo AgainSaveAtt
'      End If
'   End If
'   If strRFile <> "" Then
'      MsgBox strRFile & vbCrLf & "上列電子檔已存在系統中，請查明！", vbExclamation, "發生錯誤！"
'   End If
'
'   Exit Sub
'
'CheckingErr:
'   cnnConnection.RollbackTrans
'   If Err.Description <> "" Then MsgBox (Err.Description)
'End Sub

'Add By Sindy 2015/3/10 加入回覆單
'Private Function OpenAddAtt() As Boolean
'   Dim stFileName As String
'   Dim sFile
'   Dim ii As Integer
'   Dim strFile As String
'   Dim strFSize As String, strDateLastModified As String, strCaseNo As String
'
''on error GoTo ErrHnd
'
'   OpenAddAtt = True
'   stFileName = "*.PDF"
'   With CommonDialog1
'      .CancelError = True
'      .FileName = stFileName
'      .Filter = "All Files (*.PDF)|*.PDF"
'      If GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "") <> "" Then
'         .InitDir = GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "")
'      Else
'         .InitDir = PUB_Getdesktop
'      End If
'      .MaxFileSize = 3000
'      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
'      .ShowOpen
'      If .FileName <> "" Then
'         If InStr(.FileName, ChrW$(0)) > 0 Then
'            sFile = Split(.FileName, ChrW$(0))
'            '記錄路徑
'            SaveSetting "TAIE", "FCP", EMP_回覆單 & "Dir", sFile(0)
'            For ii = 1 To UBound(sFile)
'               If InStr(sFile(ii), "\") > 0 Then
'                  stFileName = sFile(ii)
'               Else
'                  stFileName = sFile(0) & "\" & sFile(ii)
'               End If
'               If ChkAddFile(sFile(ii), stFileName, strFSize, strDateLastModified, strCaseNo) = False Then
'                  OpenAddAtt = False
'                  Exit Function
'               End If
'               AddListX stFileName & " (" & Round(strFSize / 1024, 2) & " KB)" & " #" & Format(strDateLastModified, "YYYYMMDDHHMMSS"), strCaseNo
'            Next ii
'         Else
'            strFile = Mid(.FileName, InStrRev(.FileName, "\") + 1)
'            '記錄路徑
'            If InStr(.FileName, "\") > 0 Then
'               For ii = Len(.FileName) To 1 Step -1
'                  If Mid(Trim(.FileName), ii, 1) = "\" Then
'                     SaveSetting "TAIE", "FCP", EMP_回覆單 & "Dir", Mid(Trim(.FileName), 1, ii - 1)
'                     Exit For
'                  End If
'               Next ii
'            End If
'            stFileName = .FileName
'            If ChkAddFile(strFile, stFileName, strFSize, strDateLastModified, strCaseNo) = False Then
'               OpenAddAtt = False
'               Exit Function
'            End If
'            AddListX stFileName & " (" & Round(strFSize / 1024, 2) & " KB)" & " #" & Format(strDateLastModified, "YYYYMMDDHHMMSS"), strCaseNo
'         End If
'      End If
'   End With
'   Exit Function
'
'ErrHnd:
'   If Err.Number <> 32755 Then
'      MsgBox Err.Description
'   Else
'      OpenAddAtt = False '取消
'   End If
'End Function

'Add By Sindy 2015/3/12 檢查使用者點選的檔案
'Private Function ChkAddFile(ByVal strChkFile As String, ByVal strFullFile As String, _
'                            ByRef strFSize As String, ByRef strDateLastModified As String, _
'                            ByRef strCaseNo As String) As Boolean
'Dim bolChkFileOK As Boolean
'Dim fs, f
''Dim strCP01 As String, strCP02 As String, strCP0304 As String
'
'   If InStr(strChkFile, "#") > 0 Or InStr(strChkFile, "&") > 0 Then
'      MsgBox strChkFile & vbCrLf & vbCrLf & "【#和&】符號為系統保留字，不可使用於檔案命名！", vbExclamation
'      ChkAddFile = False
'      Exit Function
'   End If
'
'   '解析電子檔的本所案號
''   strCP01 = ""
''   strCP02 = ""
''   strCP0304 = ""
''   strCaseNo = Left(strChkFile, InStr(strChkFile, ".") - 1)
''   If InStr(strCaseNo, "-") > 0 Then
''      strCP0304 = Mid(strCaseNo, InStr(strCaseNo, "-"))
''      strCaseNo = Left(strCaseNo, InStr(strCaseNo, "-") - 1)
''   End If
''   For i = 1 To Len(strCaseNo)
''      If (Asc(Mid(strCaseNo, i, 1)) >= 65 And Asc(Mid(strCaseNo, i, 1)) <= 90) Or _
''         Asc(Mid(strCaseNo, i, 1)) >= 97 And Asc(Mid(strCaseNo, i, 1)) <= 122 Then
''         strCP01 = strCP01 & UCase(Mid(strCaseNo, i, 1))
''      Else
''         strCP02 = strCP02 & Mid(strCaseNo, i, 1)
''      End If
''   Next i
''   strCP02 = Format(strCP02, "000000")
''   strCaseNo = strCP01 & "-" & strCP02 & IIf(strCP0304 = "", "-0-00", strCP0304)
'   'Modify By Sindy 2015/5/28
'   strCaseNo = PUB_AnalysisFileNmGetCaseNO(strChkFile)
'   '2015/5/28 END
'   If InStr(StrCompCp01020304, strCaseNo) = 0 Then
'      MsgBox strChkFile & vbCrLf & vbCrLf & "檔案有誤，無此案件!!!"
'      ChkAddFile = False
'      Exit Function
'   End If
'
'   '檢查檔名規則
'   TmpArrCaseNo = Split(StrCompCp01020304, ",")
'   bolChkFileOK = False
'   For i = 0 To UBound(TmpArrCaseNo) - 1
'      If PUB_ChkEmpFlowFNMRule(CStr(TmpArrCaseNo(i)), strChkFile, EMP_會修, "", , , , False) = True Then
'         bolChkFileOK = True
'         Exit For
'      End If
'   Next i
'   If bolChkFileOK = False Then
'      MsgBox strChkFile & vbCrLf & vbCrLf & "檔案命名不符規定，請修改檔名!!!"
'      ChkAddFile = False
'      Exit Function
'   End If
'
'   '只可加入PDF檔
'   If UCase(Mid(strChkFile, InStrRev(strChkFile, ".") + 1)) <> "PDF" Then
'      MsgBox strChkFile & vbCrLf & vbCrLf & "只可加入PDF檔！", vbExclamation
'      ChkAddFile = False
'      Exit Function
'   End If
'   '檢查檔案是否正在使用中
'   If PUB_ChkFileOpening(strFullFile) = True Then
'      MsgBox strFullFile & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation
'      ChkAddFile = False
'      Exit Function
'   End If
'   Set fs = CreateObject("Scripting.FileSystemObject")
'   Set f = fs.GetFile(strFullFile)
'   '檔案大小為 0 KB 有誤
'   If f.Size = 0 Then
'      ShowMsg strChkFile & MsgText(9221)
'      ChkAddFile = False
'      Exit Function
'   '控制可以放5M以下的檔案
'   ElseIf f.Size > 5242880 Then
'      MsgBox strChkFile & vbCrLf & vbCrLf & "檔案大小不可超過 5MB！", vbExclamation
'      ChkAddFile = False
'      Exit Function
'   End If
'   strFSize = f.Size
'   strDateLastModified = f.DateLastModified
'
'   ChkAddFile = True
'End Function

'Add By Sindy 2015/3/10
'Private Function AddListX(stNewItem As String, strCaseNo As String) As Boolean
'   Dim idx As Integer, bFound As Boolean, stFileName As String
'
'   If stNewItem <> "" Then
'      For idx = 1 To GrdAtt.Rows - 1
'         stFileName = GetFileName(GrdAtt.TextMatrix(idx, 1))
'         If UCase(GetFileName(stNewItem)) = UCase(stFileName) Then
'            MsgBox "附件 " & stFileName & " 已存在！", vbExclamation
'            AddListX = False
'            bFound = True
'            Exit For
'         End If
'      Next
'      If bFound = False Then
'         If GrdAtt.Rows = 2 Then
'            If GrdAtt.TextMatrix(1, 0) = "" Then
'               GrdAtt.TextMatrix(1, 0) = stNewItem
'               GrdAtt.TextMatrix(1, 1) = GetFileName(stNewItem)
'               GrdAtt.TextMatrix(1, 2) = strCaseNo
'            Else
'               GrdAtt.AddItem stNewItem
'               GrdAtt.TextMatrix(GrdAtt.Rows - 1, 1) = GetFileName(stNewItem)
'               GrdAtt.TextMatrix(GrdAtt.Rows - 1, 2) = strCaseNo
'            End If
'         Else
'            GrdAtt.AddItem stNewItem
'            GrdAtt.TextMatrix(GrdAtt.Rows - 1, 1) = GetFileName(stNewItem)
'            GrdAtt.TextMatrix(GrdAtt.Rows - 1, 2) = strCaseNo
'         End If
'         AddListX = True
'      End If
'   End If
'End Function
'end 2018/05/24

'2011/8/9 add by sonia 還原顏色
'Add By Sindy 2014/6/3 +intRow As Integer
Private Sub recovercolor(intRow As Integer)
Dim strCP06 As String, strCP07 As String
   
   'Modify By Sindy 2012/8/24
   'grdDataList.col = 6
   grdDataList.row = intRow 'Add By Sindy 2014/6/3
   'Modified by Lydia 2021/05/10 改成變數
   'grdDataList.col = 7 '本所期限
   grdDataList.col = colCP06 '本所期限
   '2012/8/24 End
   strCP06 = Mid(grdDataList.Text, 2)
   'Modified by Lydia 2021/05/10 改成變數
   'grdDataList.col = 8 '法定期限
   grdDataList.col = colCP06 + 1 '法定期限
   strCP07 = grdDataList.Text
   'Modify By Sindy 2023/2/8 紅色：當天期限或逾期
   'If strCP06 = ChangeTStringToTDateString(strSrvDate(2)) Then
   If strCP06 <= ChangeTStringToTDateString(strSrvDate(2)) Then
   '2023/2/8 END
      For j = 0 To grdDataList.Cols - 1
         grdDataList.col = j
         grdDataList.CellBackColor = &H8080FF
      Next j
   Else
      'Modify By Sindy 2012/8/24
      'grdDataList.col = 8
      'Modified by Lydia 2021/05/10 改成變數
      'grdDataList.col = 9 '本所案號
      grdDataList.col = colCaseNo
      '2012/8/24 End
      If UCase(Mid(grdDataList.Text, 1, 3)) = "CFT" Then
         'Modify By Sindy 2012/8/24
         'grdDataList.col = 11
         'Modified by Lydia 2021/05/10 改成變數
         'grdDataList.col = 12 '案件性質
         ''2012/8/24 End
         grdDataList.col = colCPM
         If Trim(grdDataList.Text) = "延展" Or Trim(grdDataList.Text) = "使用宣誓" Then
            For j = 0 To grdDataList.Cols - 1
               grdDataList.col = j
               grdDataList.CellBackColor = &HC000&
            Next j
         End If
      Else
         For j = 0 To grdDataList.Cols - 1
            grdDataList.col = j
            grdDataList.CellBackColor = &H80000018
         Next j
      End If
   End If
   'Add By Sindy 2014/6/25
   If Check1.Value = 1 And Check5.Value = 1 Then '有2.未處理 及 含此期間法定期限案件(逾本所期限)
      'Modified by Lydia 2021/05/10 改用變數
      'If (Trim(grdDataList.TextMatrix(intRow, 5)) = "未處理" Or Trim(grdDataList.TextMatrix(intRow, 5)) = "未函知" Or _
         Trim(grdDataList.TextMatrix(intRow, 5)) = "結案中") Then
      If (Trim(grdDataList.TextMatrix(intRow, colType)) = "未處理" Or Trim(grdDataList.TextMatrix(intRow, colType)) = "未函知" Or _
         Trim(grdDataList.TextMatrix(intRow, colType)) = "結案中") Then
         If Val(DBDATE(strCP06)) < Val(strSrvDate(1)) And _
            Val(DBDATE(strCP07)) >= Val(txtCloseDate(0) + 19110000) And _
            Val(DBDATE(strCP07)) <= Val(txtCloseDate(1) + 19110000) Then
            For j = 0 To grdDataList.Cols - 1
               grdDataList.col = j
               grdDataList.CellBackColor = &H8080FF '&HC0C0&
            Next j
         End If
      End If
   End If
   '2014/6/25 END
End Sub

'add by nickc 2006/06/23 抓離職智權人員及虛建智權人員
'2007/10/17 MODIFY BY SONIA 王協理71011同時看96030巨京專利,葉經理67002同時看96029巨京商標
'Remove by Lydia 2017/07/24 改成共用模組,已不使用
'Function GetNotInOfficeAndFalseStaff(oStr As String, oStr2 As String) As String
'Dim rsTmp2 As New ADODB.Recordset
'Dim sqlTmp2 As String
'
'   GetNotInOfficeAndFalseStaff = ""
'   sqlTmp2 = "select st01 from staff where st15>='" & oStr & "' and st15<='" & oStr2 & "' and st04='2' "
'   sqlTmp2 = sqlTmp2 & "union select st01 from staff where st15>='" & oStr & "' and st15<='" & oStr2 & "' and st04='1' and st01<'6' "
'   Select Case strUserNum
'      Case "71011"  '王協理
'         'edit by nickc 2008/04/24
'         'sqlTmp2 = sqlTmp2 & "union select st01 from staff where st15>='" & oStr & "' and st15<='" & oStr2 & "' and st04='1' and st01='96030' "
'         sqlTmp2 = sqlTmp2 & "union select st01 from staff where st15>='" & oStr & "' and st15<='" & oStr2 & "' and st04='1' and st01 in ('96031','96032') "
'
'      Case "67002" '葉經理
'         'edit by nickc 2008/04/24
'         'sqlTmp2 = sqlTmp2 & "union select st01 from staff where st15>='" & oStr & "' and st15<='" & oStr2 & "' and st04='1' and st01='96029' "
'         sqlTmp2 = sqlTmp2 & "union select st01 from staff where st15>='" & oStr & "' and st15<='" & oStr2 & "' and st04='1' and st01 in ('96029','96030') "
'      Case Else
'   End Select
'   Set rsTmp2 = New ADODB.Recordset
'   With rsTmp2
'       .CursorLocation = adUseClient
'       .Open sqlTmp2, cnnConnection, adOpenStatic, adLockReadOnly
'       If .RecordCount <> 0 Then
'           .MoveFirst
'           Do While Not .EOF
'               GetNotInOfficeAndFalseStaff = GetNotInOfficeAndFalseStaff & "'" & CheckStr(.Fields(0)) & "',"
'               .MoveNext
'           Loop
'       End If
'   End With
'   Set rsTmp2 = Nothing
'End Function
'end 2017/07/24

Private Sub txtZone_LostFocus()
   'add by nickc 2006/11/21
'   If Mid(txtSalesArea, 1, 1) <> "S" And Mid(txtSalesArea1, 1, 1) <> "S" Then
'       chkNP.Enabled = False
'       chkNP.Value = 0
'   Else
'       chkNP.Enabled = True
'   End If
End Sub

Private Sub txtZone_Validate(Cancel As Boolean)
   'add by nickc 2006/11/21
'   If Mid(txtSalesArea, 1, 1) <> "S" And Mid(txtSalesArea1, 1, 1) <> "S" Then
'       chkNP.Enabled = False
'       chkNP.Value = 0
'   Else
'       chkNP.Enabled = True
'   End If
End Sub

'Add By Sindy 2010/6/17
Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Me.Text1(Index)
   CloseIme
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'Added by Lydia 2015/06/18 +客戶中文名稱查詢
Private Sub txtCuName_GotFocus()
   'Added by Lydia 2015/06/25 輸入客戶名稱按Enter自動執行搜尋後,執行查詢
   cmdSearch.Default = False: cmdFind.Default = True
   
   TextInverse Me.txtCuName
   OpenIme
End Sub

'Added by Lydia 2015/06/25 輸入客戶名稱按Enter自動執行搜尋後,執行查詢
Private Sub txtCuName_LostFocus()
   cmdSearch.Default = True: cmdFind.Default = False
End Sub
'Added by Lydia 2015/06/18 +客戶中文名稱查詢
Private Sub cmdFind_Click()
   If Me.txtCuName.Text = "" Then
      MsgBox "請輸入客戶中文名稱的關鍵字!!!", vbExclamation + vbOKOnly
      Me.txtCuName.SetFocus
      txtCuName_GotFocus
      Exit Sub
   End If
   frm090801_1.m_strCustChnName = Me.txtCuName.Text
   frm090801_1.lblName.Caption = Me.txtCuName.Text
   m_blnOneRec = False
   m_strCustCode = ""
   If frm090801_1.StrMenu = True Then
      If frm090801_1.m_blnOneRec = False Then
         frm090801_1.Show vbModal
      End If
      m_blnOneRec = frm090801_1.m_blnOneRec
      m_strCustCode = frm090801_1.m_strCustCode
      Unload frm090801_1
   Else
      Unload frm090801_1
   End If
   If m_blnOneRec = True And m_strCustCode <> "" Then
      Me.txtCU1.Text = m_strCustCode
      Me.txtCU2.Text = IIf(Right(m_strCustCode, 3) = "000", Mid(m_strCustCode, 1, 6) & "ZZZ", IIf(Right(m_strCustCode, 1) = "0", Mid(m_strCustCode, 1, 8) & "Z", m_strCustCode))
      Me.txtCuName.Text = GetCustomerName(m_strCustCode)
   End If
   'Added by Lydia 2015/06/25 輸入客戶名稱按Enter自動執行搜尋後,執行查詢
   If Me.txtCU1.Text <> "" And Me.txtCU2.Text <> "" Then
      Call cmdSearch_Click
   End If
End Sub
'Added by Morgan 2016/3/1
Private Function GetValue(pRow As Integer, pFieldName As String) As String
   Dim iCol As Integer
   With grdDataList
   For iCol = 0 To .Cols - 1
      If UCase(.TextMatrix(0, iCol)) = UCase(pFieldName) Then
         GetValue = .TextMatrix(pRow, iCol)
         Exit For
      End If
   Next
   End With
End Function

Private Sub KillTemp()
On Error GoTo ErrHnd
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\*.*"
   End If
   Exit Sub
   
ErrHnd:
   Resume Next
End Sub

'Add by Amy 2018/05/24 frm100123_1回傳後設定Grd
Public Sub SetGrdColor(ByVal stErr As String, bolPreForm As Boolean)
    Dim j As Integer, k As Integer
    Dim strMsg As String
    Dim bolSetColor As Boolean
    Dim arrTmp
    
    'frm100123_1 按「確定」
    If bolPreForm = True Then
        grdDataList.MousePointer = flexDefault
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If stErr <> MsgText(601) Then arrTmp = Split(stErr, ",")
    For i = 1 To grdDataList.Rows - 1
        grdDataList.col = 0
        grdDataList.row = i
        If Trim(grdDataList.Text) = "V" Then
            bolSetColor = True
            If stErr <> MsgText(601) Then
                If UBound(arrTmp) = 0 Then
                    If InStr(arrTmp(0), GetValue(i, "本所案號")) > 0 Then
                        '錯誤的本所案號勾選及顏色保留
                        bolSetColor = False
                    End If
                Else
                    For k = LBound(arrTmp) To k = UBound(arrTmp)
                        If InStr(arrTmp(k), GetValue(i, "本所案號")) > 0 Then
                            '錯誤的本所案號勾選及顏色保留
                            bolSetColor = False
                            Exit For
                        End If
                    Next k
                End If
            End If
            If bolSetColor = True Then
                'Modified by Lydia 2021/05/10 改用變數
                'grdDataList.col = 7 '本所期限
                'grdDataList.Text = "*" & grdDataList.TextMatrix(i, 7)
                grdDataList.col = colCP06 '本所期限
                grdDataList.Text = "*" & grdDataList.TextMatrix(i, colCP06)
                'end 2021/05/10
                grdDataList.col = 0
                grdDataList.Text = "" '設不勾選
                For j = 0 To grdDataList.Cols - 1
                    grdDataList.col = j
                    If grdDataList.CellBackColor = &HFFC0C0 Then
                        grdDataList.CellBackColor = &H80000018 '反白
                    Else
                        grdDataList.CellBackColor = &H8080FF '變紅
                    End If
                Next j
            End If
        End If
    Next i
    If stErr <> MsgText(601) Then
        For i = LBound(arrTmp) To UBound(arrTmp)
            strMsg = strMsg & Replace(arrTmp(i), "@@", "：") & IIf(i <> UBound(arrTmp), vbCrLf, "")
        Next i
    End If
    If strMsg <> MsgText(601) Then MsgBox strMsg, vbCritical
    grdDataList.MousePointer = flexDefault
    Screen.MousePointer = vbDefault
End Sub

'產生Excel 及PDF 寄出
Private Function SaveExcel() As Boolean
    Dim xlsAgentPoint As New Excel.Application
    Dim Wks As New Worksheet
    Dim arrTmp As Variant, arrWidth As Variant
    Dim i As Integer, j As Integer, iCol As Integer, intField As Integer, intCounter As Integer
    Dim strAllField As String, strTmp As String, strTmpW As String
    
On Error GoTo ErrHand
    SaveExcel = False
    intField = 65: intCounter = 1
    xlsFileName = GetDepartmentName(txtSalesArea) & " " & GetDeptMan(txtSalesArea) & " 期限資料-" & Format(Now, "yyyymmddhhmmss") & MsgText(43)
    If Dir(strExcelPath & xlsFileName) = MsgText(601) Then
        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
        End If
    Else
            Kill strExcelPath & xlsFileName
    End If
            
    xlsAgentPoint.SheetsInNewWorkbook = 3 '預設工作表數量
    xlsAgentPoint.Workbooks.add
    xlsAgentPoint.Application.WindowState = xlMinimized
    Set Wks = xlsAgentPoint.Worksheets(1)
    
    For i = 0 To grdDataList.Rows - 1
        iCol = 0
        For j = 1 To grdDataList.Cols - 1
            If grdDataList.ColWidth(j) > 0 Then
                '欄位名稱
                If i = 0 Then
                    strTmp = strTmp & "," & grdDataList.TextMatrix(i, j)
                    Select Case grdDataList.TextMatrix(i, j)
                        Case "所別"
                            strTmpW = strTmpW & "," & 4.7
                        Case "本所案號"
                            strTmpW = strTmpW & "," & 14.38
                        Case "案件名稱"
                            strTmpW = strTmpW & "," & 12
                        Case "本所案號"
                            strTmpW = strTmpW & "," & 14.38
                        Case "申請人"
                            strTmpW = strTmpW & "," & 11
                        Case "申請案號"
                            strTmpW = strTmpW & "," & 15
                        Case "收文號"
                            strTmpW = strTmpW & "," & 9.5
                        Case Else
                            strTmpW = strTmpW & "," & 8.38
                    End Select
                '資料
                Else
                    arrTmp(iCol) = grdDataList.TextMatrix(i, j)
                    iCol = iCol + 1
                End If
            End If
        Next j
        If i = 0 Then
            arrTmp = Split(Mid(strTmp, 2), ",")
        End If
        Wks.Range(Chr(intField) & intCounter & ":" & Chr(intField + UBound(arrTmp)) & intCounter).Value = arrTmp
        intCounter = intCounter + 1
    Next i
    
    '欄寬
    arrWidth = Split(Mid(strTmpW, 2), ",")
    For i = LBound(arrWidth) To UBound(arrWidth)
        Wks.Columns(Chr(i + intField) & ":" & Chr(i + intField)).ColumnWidth = arrWidth(i)
        If Val(arrWidth(i)) > 8.38 Then
            Wks.Columns(Chr(i + intField) & ":" & Chr(i + intField)).WrapText = True
        End If
    Next i
    
    Wks.Range(Chr(intField) & "1:" & Chr(intField + UBound(arrTmp)) & intCounter - 1).Font.Size = 11
    '框線
    Wks.Range(Chr(intField) & "1:" & Chr(intField + UBound(arrTmp)) & intCounter - 1).Select
    xlsAgentPoint.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    xlsAgentPoint.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    xlsAgentPoint.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    xlsAgentPoint.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    xlsAgentPoint.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    xlsAgentPoint.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    
    Wks.PageSetup.PaperSize = 9 'A4
    Wks.PageSetup.Orientation = xlLandscape '橫印
    Wks.PageSetup.PrintTitleRows = "$1:$1"  '表頭保留
    Wks.PageSetup.LeftMargin = xlsAgentPoint.InchesToPoints(0.5)
    Wks.PageSetup.RightMargin = xlsAgentPoint.InchesToPoints(0.5)
    Wks.PageSetup.TopMargin = xlsAgentPoint.InchesToPoints(0.2)
    Wks.PageSetup.BottomMargin = xlsAgentPoint.InchesToPoints(0.2)
    Wks.PageSetup.HeaderMargin = xlsAgentPoint.InchesToPoints(0.3)
    Wks.PageSetup.FooterMargin = xlsAgentPoint.InchesToPoints(0.3)
    '設定一頁列印
    Wks.PageSetup.PrintArea = "$A$1:$" & Chr(intField + UBound(arrTmp)) & intCounter - 1
    xlsAgentPoint.ActiveWindow.View = xlPageBreakPreview
    Wks.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
    
    If Val(xlsAgentPoint.Version) < 12 Then
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
    Else
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
     End If
    xlsAgentPoint.Workbooks.Close
    xlsAgentPoint.Quit
    Set xlsAgentPoint = Nothing
    Set Wks = Nothing
    SaveExcel = True
    Exit Function
ErrHand:
    MsgBox Err.Description, , MsgText(5)
    If Val(xlsAgentPoint.Version) < 12 Then
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
    Else
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
    End If
    xlsAgentPoint.Workbooks.Close
    xlsAgentPoint.Quit
    Set xlsAgentPoint = Nothing
    Set Wks = Nothing
End Function
'end 2020/03/25

'Added by Lydia 2021/05/10 增加管制備註功能：調整版面
Private Function doQuery_1() As Boolean
Dim stCon As String, stConST As String
Dim stCon1 As String, stCon2 As String, stCon3 As String, stCon4 As String
Dim stCon5 As String, stCon6 As String, strCP13 As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strInData As String
'Dim stVTBw As String 'Add by Morgan 2010/1/27 預定收款日資料
Dim stIdList As String, stConId As String 'Add by Morgan 2010/1/29
Dim bolChgSystemkind As Boolean, strOldSystemkind As String
Dim stCon8 As String 'Add by Morgan 2011/8/15
Dim stCon9 As String 'Add by Morgan 2012/8/21
Dim strChkSql As String 'Add By Sindy 2014/6/12
Dim strMidSql As String
Dim stConP As String, stConT As String, stConS As String, stConL As String, stConH As String

   m_blnColOrderAsc = True 'Add By Sindy 2014/6/19
   LblCntTime.Caption = "執行時間：" & Format(ServerTime, "##:##:##")
   
   stConP = "": stConT = "": stConS = "": stConL = "": stConH = ""
   
   'add by nickc 2007/01/23
   Dim stCon1_1 As String, stCon2_1 As String
   stCon1_1 = "": stCon2_1 = ""

   stCon4 = ""

   stCon5 = "" 'Add By Sindy 2011/6/21 未續簽
   stCon6 = "" 'Add By Sindy 2011/6/21 未回執
   
   stCon8 = "" '未收款
   
   stCon = "": stConST = "": stCon1 = "": stCon2 = ""
   stCon3 = "" '查詢行事曆資料

   bolChgSystemkind = False
   If Text1(0) = "" Or Text1(1) = "" Then
      If systemkind = "" Then
          systemkind = "ALL"
      End If
   Else
      bolChgSystemkind = True
      strOldSystemkind = systemkind
      systemkind = Text1(0).Text
   End If
   
   '所別
   '林柄佑要控制所別
   If strUserNum = "82026" Then
      stConST = stConST & " and s2.st06 = '" & pub_strUserOffice & "'"
      stCon5 = stCon5 & " and s2.st06 = '" & pub_strUserOffice & "'"
      stCon6 = stCon6 & " and s2.st06 = '" & pub_strUserOffice & "'"
   End If
   
   '陳經理查詢所有智權人員要控制系統類別
   If strUserNum = "68005" And txtSales <> "68005" Then
      systemkind = "CFT,FCT,S,CFC"
   End If
   
   '區別
   'Modify by Amy 2019/02/12 總經理業務工作代理人員
   If bolSpecMan = True And (InStr(strSpecCode, "A8") > 0 Or InStr(strSpecCode, "總經理業務工作代理人員") > 0) And txtSales <> strUserNum Then
       '開放專利處部份智權同仁資料給彥葶代為處理,不考慮業務區(因彥葶與開放的智權同仁業務區不同)
   ElseIf txtSales = "80030" Or txtSales = "79075" Or _
      (Trim(txtSales) <> "" And PUB_GetST05Limits(strUserNum) = True And txtSales.Enabled = True And txtSales <> strUserNum) Then
      '不限制區別
  
   '查自己資料不限制區別,因為有調區問題
   ElseIf txtSales = strUserNum Then
   Else
         If txtSalesArea <> "" Then
            stCon1 = stCon1 & " and s1.st15>='" & txtSalesArea & "'"
            stCon2 = stCon2 & " and s2.st15||''>='" & txtSalesArea & "'" 'cp12
            stCon3 = stCon3 & " and s2.st15>='" & txtSalesArea & "'"
            'add by nickc 2008/04/24 加入未收款
            stCon4 = stCon4 & " and a0k22||''>='" & txtSalesArea & "'"
            stCon5 = stCon5 & " and s2.st15||''>='" & txtSalesArea & "'" 'cp12
            stCon6 = stCon6 & " and s2.st15||''>='" & txtSalesArea & "'" 'cp12
         End If
         If txtSalesArea1 <> "" Then
            stCon1 = stCon1 & " and s1.st15<='" & txtSalesArea1 & "'"
            stCon2 = stCon2 & " and s2.st15||''<='" & txtSalesArea1 & "'" 'cp12
            stCon3 = stCon3 & " and s2.st15<='" & txtSalesArea1 & "'"
            stCon4 = stCon4 & " and a0k22||''<='" & txtSalesArea1 & "'"
            stCon5 = stCon5 & " and s2.st15||''<='" & txtSalesArea1 & "'" 'cp12
            stCon6 = stCon6 & " and s2.st15||''<='" & txtSalesArea1 & "'" 'cp12
         End If
   End If
   
   '智權人員
   If Trim(txtSales) <> "" Then
        If (strUserNum <> "80030" And txtSales <> "80030") Then
            If bolSpecMan = True And (InStr(strSpecCode, "A8") > 0 Or InStr(strSpecCode, "總經理業務工作代理人員") > 0) And txtSales <> strUserNum Then
                '開放專利處部份智權同仁資料給彥葶代為處理,不考慮業務區(因彥葶與開放的智權同仁業務區不同)
                stIdList = PUB_GetSalesList(Trim(txtSales))
            Else
                '若不是多員工編號時用 = 算符來加速查詢
                'Modify By Sindy 2021/9/17 txtZone ==> IIf(Trim(txtSales) <> "", PUB_GetST06(Trim(txtSales)), "")
                'stIdList = PUB_GetSalesList(Trim(txtSales), txtSalesArea, txtSalesArea1, txtZone)
                stIdList = PUB_GetSalesList(Trim(txtSales), txtSalesArea, txtSalesArea1, IIf(Trim(txtSales) <> "", PUB_GetST06(Trim(txtSales)), ""))
            End If
            
            If InStr(stIdList, ",") = 0 Then
               stConId = " = " & stIdList & " "
            Else
               stConId = " in (" & stIdList & " ) "
            End If
            
            '2010/5/10 add by sonia 因中所有跨區帶人故離職智權人員的帶人主管不考慮業務區條件
            If Pub_StrST52 Then
               stCon1 = "": stCon2 = "": stCon3 = "": stCon4 = ""
            End If
            '2010/5/10 end
            
            stCon1 = stCon1 & " and np10 " & stConId
            stCon2 = stCon2 & " and cp13 " & stConId
            stCon3 = stCon3 & " and ss01 " & stConId
            stCon4 = stCon4 & " and a0k20||'' " & stConId
            stCon5 = stCon5 & " and s2.st01 " & stConId 'cp13
            stCon6 = stCon6 & " and s2.st01 " & stConId 'cp13
        Else
           '查80030洪琬姿時同時查F4103
            If txtSales = "80030" Then

               StrSQLa = "select ST01 from STAFF where ST04<>'1' and ST03 like 'F1%' "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               strInData = "'80030','F4103'"
               If rsA.RecordCount > 0 Then
                  rsA.MoveFirst
                  Do While rsA.EOF = False
                     strInData = strInData & ",'" & rsA.Fields(0).Value & "'"
                     rsA.MoveNext
                  Loop
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               
               stCon1 = stCon1 & " and np10 IN (" & strInData & ") "
               stCon2 = stCon2 & " and cp13='" & Trim(txtSales) & "'"
               stCon3 = stCon3 & " and ss01='" & Trim(txtSales) & "'"
               stCon4 = stCon4 & " and a0k20||'' IN (" & strInData & ") "
               stCon5 = stCon5 & " and s2.st01='" & Trim(txtSales) & "'"
               stCon6 = stCon6 & " and s2.st01='" & Trim(txtSales) & "'"
            Else
               stCon1 = stCon1 & " and np10='" & Trim(txtSales) & "' "
               stCon2 = stCon2 & " and cp13='" & Trim(txtSales) & "'"
               stCon3 = stCon3 & " and ss01='" & Trim(txtSales) & "'"
               stCon4 = stCon4 & " and a0k20||''='" & Trim(txtSales) & "'"
               stCon5 = stCon5 & " and s2.st01='" & Trim(txtSales) & "'"
               stCon6 = stCon6 & " and s2.st01='" & Trim(txtSales) & "'"
            End If
        End If
   '智權人員 為空
   Else
        If bolSpecMan = True And InStr(strSpecCode, "A8") > 0 Then
            'A2023彥葶登入,未輸智權人員-設定查A7人員
            stConId = " in ('" & Replace(Pub_GetSpecMan("A7"), ";", "','") & "')"
            stCon1 = stCon1 & " and np10 " & stConId
            stCon2 = stCon2 & " and cp13 " & stConId
            stCon3 = stCon3 & " and ss01 " & stConId
            '加入未收款
            stCon4 = stCon4 & " and a0k20||'' " & stConId
            stCon5 = stCon5 & " and s2.st01 " & stConId
            stCon6 = stCon6 & " and s2.st01 " & stConId
        End If
   End If
   
   stCon8 = stCon2 'Add by Morgan 2011/8/15 未收款(新)

   stCon1_1 = stCon1
   stCon2_1 = stCon2
   stCon9 = stCon1
   
   '本所期限
   If txtCloseDate(0) <> "" Then
      If Check5.Value = 0 Then
         stCon1 = stCon1 & " and np08>=" & ChangeTStringToWString(txtCloseDate(0))
      End If
      stCon9 = stCon9 & " and np23>=" & ChangeTStringToWString(txtCloseDate(0)) 'Added by Morgan 2012/8/21
      
      stCon2 = stCon2 & " and cp06>=" & ChangeTStringToWString(txtCloseDate(0))
      stCon3 = stCon3 & " and ss02>=" & ChangeTStringToWString(txtCloseDate(0))
      stCon1_1 = stCon1_1 & " and np09>=" & ChangeWDateStringToWString(DateAdd("d", 10, ChangeWStringToWDateString(ChangeTStringToWString(txtCloseDate(0)))))
      stCon2_1 = stCon2_1 & " and cp07>=" & ChangeWDateStringToWString(DateAdd("d", 10, ChangeWStringToWDateString(ChangeTStringToWString(txtCloseDate(0)))))
      stCon5 = stCon5 & " and cp54>=" & ChangeTStringToWString(txtCloseDate(0))
      '發文日+7天(日曆天)符合查詢本所期限條件
      '2011/7/6 MODIFY BY SONIA 起日抓三個月內
      stCon6 = stCon6 & " and cp27>=" & ChangeWDateStringToWString(DateAdd("m", -3, ChangeWStringToWDateString(ChangeTStringToWString(txtCloseDate(0)))))
   End If
   If txtCloseDate(1) <> "" Then
      If Check5.Value = 0 Then
         stCon1 = stCon1 & " and np08<=" & ChangeTStringToWString(txtCloseDate(1))
      End If
      stCon9 = stCon9 & " and np23<=" & ChangeTStringToWString(txtCloseDate(1))
      
      stCon2 = stCon2 & " and cp06<=" & ChangeTStringToWString(txtCloseDate(1))
      stCon3 = stCon3 & " and ss02<=" & ChangeTStringToWString(txtCloseDate(1))
      stCon1_1 = stCon1_1 & " and np09<=" & ChangeWDateStringToWString(DateAdd("d", 10, ChangeWStringToWDateString(ChangeTStringToWString(txtCloseDate(1)))))
      stCon2_1 = stCon2_1 & " and cp07<=" & ChangeWDateStringToWString(DateAdd("d", 10, ChangeWStringToWDateString(ChangeTStringToWString(txtCloseDate(1)))))
      stCon5 = stCon5 & " and cp54<=" & ChangeTStringToWString(txtCloseDate(1))
      '發文日+7天(日曆天)符合查詢本所期限條件
      stCon6 = stCon6 & " and cp27<=" & ChangeWDateStringToWString(DateAdd("d", -7, ChangeWStringToWDateString(ChangeTStringToWString(txtCloseDate(1)))))
   End If
   
   '含此期間法定期限案件(逾本所期限)
   If Check5.Value = 1 Then
      stCon1 = stCon1 & " and ((np08>=" & ChangeTStringToWString(txtCloseDate(0)) & " And np08<=" & ChangeTStringToWString(txtCloseDate(1)) & ") or (np08<" & strSrvDate(1) & " and np09>=" & ChangeTStringToWString(txtCloseDate(0)) & " And np09 <=" & ChangeTStringToWString(txtCloseDate(1)) & "))"
   End If
   
   '本所案號
   If Text1(0) <> "" And Text1(1) <> "" Then
      stCon1 = stCon1 & " and (np02='" & Text1(0) & "' and np03='" & Text1(1) & "' and np04='" & Text1(2) & "' and np05='" & Text1(3) & "') "
      stCon9 = stCon9 & " and (np02='" & Text1(0) & "' and np03='" & Text1(1) & "' and np04='" & Text1(2) & "' and np05='" & Text1(3) & "') " 'Added by Morgan 2012/8/21
      
      stCon2 = stCon2 & " and (cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "') "
      stCon4 = stCon4 & " and (cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "') "
      stCon1_1 = stCon1_1 & " and (np02='" & Text1(0) & "' and np03='" & Text1(1) & "' and np04='" & Text1(2) & "' and np05='" & Text1(3) & "') "
      stCon2_1 = stCon2_1 & " and (cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "') "
      stCon5 = stCon5 & " and (cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "') "
      stCon6 = stCon6 & " and (cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "') "
      'Add by Morgan 2011/8/15 未收款(新)
      stCon8 = stCon8 & " and (cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "') "
   End If
   
   '案件性質
   If Trim(txtCP10) <> "" Then
      stCon1 = stCon1 & " and np07=" & txtCP10 & " "
      stCon9 = stCon9 & " and np07=" & txtCP10 & " "
      
      stCon2 = stCon2 & " and cp10='" & txtCP10 & "' "
      'add by nickc 2008/04/24 加入未收款
      stCon4 = stCon4 & " and cp10='" & txtCP10 & "' "
      stCon1_1 = stCon1_1 & " and np07=" & txtCP10 & " "
      stCon2_1 = stCon2_1 & " and cp10='" & txtCP10 & "' "
      stCon5 = stCon5 & " and cp10='" & txtCP10 & "' "
      stCon6 = stCon6 & " and cp10='" & txtCP10 & "' "
      'Add by Morgan 2011/8/15 未收款(新)
      stCon8 = stCon8 & " and cp10='" & txtCP10 & "' "
   End If
   
   '未收文部分取消專利處之催審411改為晚上批次通知
   stCon1 = stCon1 & " and np02||np07 not in ('P411','PS411','CFP411','CPS411') "
   
   stCon9 = stCon9 & " and np02||np07='CFP107' " 'Added by Morgan 2012/8/21

   If Trim(txtCU1) <> "" Then
       txtCU1 = Mid(txtCU1 & "000000000", 1, 9)
       txtCU2 = Mid(txtCU2 & "000000000", 1, 9)
       stConP = stConP & " and ((pa26>='" & txtCU1 & "' and pa26<='" & txtCU2 & "') or (pa27>='" & txtCU1 & "' and pa27<='" & txtCU2 & "') or (pa28>='" & txtCU1 & "' and pa28<='" & txtCU2 & "') or (pa29>='" & txtCU1 & "' and pa29<='" & txtCU2 & "') or (pa30>='" & txtCU1 & "' and pa30<='" & txtCU2 & "')) "
       stConT = stConT & " and ((tm23>='" & txtCU1 & "' and tm23<='" & txtCU2 & "') or (tm78>='" & txtCU1 & "' and tm78<='" & txtCU2 & "') or (tm79>='" & txtCU1 & "' and tm79<='" & txtCU2 & "') or (tm80>='" & txtCU1 & "' and tm80<='" & txtCU2 & "') or (tm81>='" & txtCU1 & "' and tm81<='" & txtCU2 & "')) "
       stConS = stConS & " and ((sp08>='" & txtCU1 & "' and sp08<='" & txtCU2 & "') or (sp58>='" & txtCU1 & "' and sp58<='" & txtCU2 & "') or (sp59>='" & txtCU1 & "' and sp59<='" & txtCU2 & "') or (sp65>='" & txtCU1 & "' and sp65<='" & txtCU2 & "') or (sp66>='" & txtCU1 & "' and sp66<='" & txtCU2 & "')) "
       stConL = stConL & " and ((lc11>='" & txtCU1 & "' and lc11<='" & txtCU2 & "') or (lc43>='" & txtCU1 & "' and lc43<='" & txtCU2 & "') or (lc44>='" & txtCU1 & "' and lc44<='" & txtCU2 & "') or (lc45>='" & txtCU1 & "' and lc45<='" & txtCU2 & "') or (lc46>='" & txtCU1 & "' and lc46<='" & txtCU2 & "')) "
       stConH = stConH & " and ((hc05>='" & txtCU1 & "' and hc05<='" & txtCU2 & "') or (hc24>='" & txtCU1 & "' and hc24<='" & txtCU2 & "') or (hc25>='" & txtCU1 & "' and hc25<='" & txtCU2 & "') or (hc26>='" & txtCU1 & "' and hc26<='" & txtCU2 & "') or (hc27>='" & txtCU1 & "' and hc27<='" & txtCU2 & "')) "
       stCon4 = stCon4 & " and ((A0K03>='" & txtCU1 & "' and A0K03<='" & txtCU2 & "')) "
   End If
   
On Error GoTo ErrHnd
   
'Memo by Lydia 2021/05/10
'相關Table:
'R100123: 確定刪除Table
'R100123_1：期限資料查詢(frm100123) 增加管制備註功能－調整後使用 ==>使用中
'R100123_2: 期限資料查詢 (frm100123) - 調整前使用 ==>無使用了
'end 2021/05/10

'暫存檔規格
'R100123_1:
'   RCP01   VARCHAR2(3)    NULL,
'   RCP02   VARCHAR2(6)    NULL,
'   RCP03   VARCHAR2(1)    NULL,
'   RCP04   VARCHAR2(2)    NULL,
'   RCP09   VARCHAR2(9)    NULL,
'   RCP12   VARCHAR2(3)    NULL,       管制部門
'   RCP13   VARCHAR2(6)    NULL,
'   RKIND   VARCHAR2(2)    NULL,       分類
'   ID      VARCHAR2(6)    NULL,       strUserNum
'   RCP06   VARCHAR2(12)   NULL,       本所期限
'   REMP    VARCHAR2(6)    NULL,       管制人
'   RCP07   VARCHAR2(12)   NULL,       法定期限
'   RNP23   VARCHAR2(12)   NULL,       約定期限
'   RSUBNO  VARCHAR2(50)   NULL,       分所號
'   RCASENAME  VARCHAR2(180)     NULL, 案件名稱
'   RCP10   VARCHAR2(4)    NULL,
'   RCP14   VARCHAR2(6)    NULL,
'   RCP05   VARCHAR2(12)   NULL,
'   RCP27   VARCHAR2(12)   NULL,
'   RAPPID  VARCHAR2(9)    NULL,       申請人1
'   RNATION VARCHAR2(3)    NULL,       申請國家
'   RCASENO VARCHAR2(30)   NULL,       申請案號
'   RNP22   VARCHAR2(10)   NULL,       NP序號
'   RPKEY   VARCHAR2(20)   NULL,       ss01||'-'||ss02||'-'||ss03
'   RCP64   VARCHAR2(500)        NULL, 案件備註/下一程序備註
'   RCTLREM   VARCHAR2(500)      NULL, 管制備註

'事件分類:'1','未發文','2','未處理','3','行事曆','4','未回覆','5','未通知'
'        ,'6','未收款','7','未續簽','8','未回執','9','未函知','10','結案中'

'***********************************
   strSql = "delete R100123_1 where id='" & strUserNum & "'"
   cnnConnection.Execute strSql, intI
   '預設Insert語法
   'Modify By Sindy 2021/9/16 + ,RECVTYPE : 類型
   strMidSql = "insert into R100123_1 (RCP01,RCP02,RCP03,RCP04,RCP09,RCP12,RCP13,RKIND,ID," & _
                     "RCP06,REMP,RCP07,RNP23,RSUBNO,RCASENAME,RCP10,RCP14,RCP05,RCP27," & _
                     "RAPPID,RNATION,RCASENO,RNP22,RPKEY,RCP64,RCTLREM,RECVTYPE) "
'***********************************
   
   '7.未續簽
   strSql = strMidSql & " select cp01,cp02,cp03,cp04,cp09,nvl(getlos04(cp09,3),s2.st15) as 管制人部門,nvl(getlos04(cp09),s2.st01),'7','" & strUserNum & "', " & _
             " ' '||sqldatet(cp54),cp14 as 管制人,sqldatet(cp07) as 法定期限,'' as 約定期限,hc07 as 分所號,HC06 As 案件名稱,cp10,cp14,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日," & _
             " hc05,'000','' as 申請案號,cp66 as 序號,'' as PKey,substr(cp64,1,500) as caserem,substr(cmr04,1,500) ctlrem,decode(substr(cp09,1,1),'B','B類','') as RECVTYPE " & _
             "from hirecase,caseprogress,staff s2, ctlremark " & _
             "where cp01||cp10='LA0' and cp158=0 and cp159=0 and substr(cp09,1,1)<>'D' " & _
             "and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) " & _
             "and cp13=s2.st01(+) and cp09=cmr01(+) and cp66=cmr02(+)" & stCon5 & stConH
   cnnConnection.Execute strSql, intI
   'Modify By Sindy 2021/7/28 抓案源介紹人 cp13=s2.st01(+) => nvl(getlos04(cp09),cp13)=s2.st01(+)
   strSql = strMidSql & " select cp01,cp02,cp03,cp04,cp09,s2.st15 as 管制人部門,s2.st01,'7','" & strUserNum & "', " & _
             " ' '||sqldatet(cp54),cp14 as 管制人,sqldatet(cp07) as 法定期限,'' as 約定期限,hc07 as 分所號,HC06 As 案件名稱,cp10,cp14,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日," & _
             " hc05,'000','' as 申請案號,cp66 as 序號,'' as PKey,substr(cp64,1,500) as caserem,substr(cmr04,1,500) ctlrem,decode(substr(cp09,1,1),'B','B類','') as RECVTYPE " & _
             "from hirecase,caseprogress,staff s2, ctlremark " & _
             "where cp01||cp10='LA0' and cp158=0 and cp159=0 and substr(cp09,1,1)<>'D' " & _
             "and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) " & _
             "and nvl(getlos04(cp09),cp13)=s2.st01(+) and cp09=cmr01(+) and cp66=cmr02(+)" & stCon5 & stConH
   cnnConnection.Execute strSql, intI
   
   '8.未回執
   strSql = strMidSql & " select cp01,cp02,cp03,cp04,cp09,nvl(getlos04(cp09,3),s2.st15) as 管制人部門," & _
             "nvl(getlos04(cp09),s2.st01),'8','" & strUserNum & "',' '||sqldatet(to_number(to_char(to_date(cp27,'YYYYMMDD')+7,'YYYYMMDD'))),cp14 as 管制人,sqldatet(cp07) as 法定期限,'' as 約定期限," & _
             "lc16 as 分所號,NVL(NVL(LC05,LC06),LC07) As 案件名稱,cp10,cp14,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,lc11,lc15,'' as 申請案號,cp66 as 序號,'' as PKey,substr(cp64,1,500) as caserem,substr(cmr04,1,500) ctlrem,decode(substr(cp09,1,1),'B','B類','') as RECVTYPE " & _
             "from lawcase,caseprogress,staff s2, ctlremark " & _
             "where cp27>=20110701 and cp01 in('L','CFL','FCL','LIN','ACS','') and cp50 is not null and cp46 is null " & _
             "and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and cp13=s2.st01(+) " & _
             "and cp09=cmr01(+) and cp66=cmr02(+) " & stCon6 & stConL
   cnnConnection.Execute strSql, intI
   'Modify By Sindy 2021/7/28 抓案源介紹人 cp13=s2.st01(+) => nvl(getlos04(cp09),cp13)=s2.st01(+)
   '" & IIf(InStr(UCase(stCon6), "CP13") > 0, "/*+ INDEX(CASEPROGRESS IDXCP13051027) */ ", "") & "
   strSql = strMidSql & " select cp01,cp02,cp03,cp04,cp09,s2.st15 as 管制人部門," & _
             "s2.st01,'8','" & strUserNum & "',' '||sqldatet(to_number(to_char(to_date(cp27,'YYYYMMDD')+7,'YYYYMMDD'))),cp14 as 管制人,sqldatet(cp07) as 法定期限,'' as 約定期限," & _
             "lc16 as 分所號,NVL(NVL(LC05,LC06),LC07) As 案件名稱,cp10,cp14,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,lc11,lc15,'' as 申請案號,cp66 as 序號,'' as PKey,substr(cp64,1,500) as caserem,substr(cmr04,1,500) ctlrem,decode(substr(cp09,1,1),'B','B類','') as RECVTYPE " & _
             "from lawcase,caseprogress,staff s2, ctlremark " & _
             "where cp27>=20110701 and cp01 in('L','CFL','FCL','LIN','ACS','') and cp50 is not null and cp46 is null " & _
             "and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and nvl(getlos04(cp09),cp13)=s2.st01(+) " & _
             "and cp09=cmr01(+) and cp66=cmr02(+) " & stCon6 & stConL
   cnnConnection.Execute strSql, intI
   
   strSql = strMidSql & " select cp01,cp02,cp03,cp04,cp09,nvl(getlos04(cp09,3),s2.st15) as 管制人部門," & _
             "nvl(getlos04(cp09),s2.st01),'8','" & strUserNum & "',' '||sqldatet(to_number(to_char(to_date(cp27,'YYYYMMDD')+7,'YYYYMMDD'))),cp14 as 管制人,sqldatet(cp07) as 法定期限,'' as 約定期限," & _
             "hc07 as 分所號,HC06 As 案件名稱,cp10,cp14,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,hc05,'000','' as 申請案號,cp66 as 序號,'' as PKey,substr(cp64,1,500) as caserem,substr(cmr04,1,500) ctlrem,decode(substr(cp09,1,1),'B','B類','') as RECVTYPE " & _
             "from hirecase,caseprogress,staff s2, ctlremark " & _
             "where cp27>=20110701 and cp01='LA' and cp50 is not null and cp46 is null " & _
             "and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and cp13=s2.st01(+) " & _
             "and cp09=cmr01(+) and cp66=cmr02(+) " & stCon6 & stConH
   cnnConnection.Execute strSql, intI
   'Modify By Sindy 2021/7/28 抓案源介紹人 cp13=s2.st01(+) => nvl(getlos04(cp09),cp13)=s2.st01(+)
   '" & IIf(InStr(UCase(stCon6), "CP13") > 0, "/*+ INDEX(CASEPROGRESS IDXCP13051027) */ ", "") & "
   strSql = strMidSql & " select cp01,cp02,cp03,cp04,cp09,s2.st15 as 管制人部門," & _
             "s2.st01,'8','" & strUserNum & "',' '||sqldatet(to_number(to_char(to_date(cp27,'YYYYMMDD')+7,'YYYYMMDD'))),cp14 as 管制人,sqldatet(cp07) as 法定期限,'' as 約定期限," & _
             "hc07 as 分所號,HC06 As 案件名稱,cp10,cp14,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日,hc05,'000','' as 申請案號,cp66 as 序號,'' as PKey,substr(cp64,1,500) as caserem,substr(cmr04,1,500) ctlrem,decode(substr(cp09,1,1),'B','B類','') as RECVTYPE " & _
             "from hirecase,caseprogress,staff s2, ctlremark " & _
             "where cp27>=20110701 and cp01='LA' and cp50 is not null and cp46 is null " & _
             "and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and nvl(getlos04(cp09),cp13)=s2.st01(+) " & _
             "and cp09=cmr01(+) and cp66=cmr02(+) " & stCon6 & stConH
   cnnConnection.Execute strSql, intI

   '2015/7/15 modify by sonia 只需管制人非查
   strSql = "SELECT * FROM R100123_1 WHERE id='" & strUserNum & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
         .MoveFirst
         Do While Not .EOF
            strCP13 = PUB_GetAKindSalesNo(.Fields("rcp01"), .Fields("rcp02"), .Fields("rcp03"), .Fields("rcp04"))
            strSql = "Update R100123_1 set rcp12='" & PUB_GetStaffST15(strCP13, "1") & "',REMP='" & strCP13 & "' where id='" & strUserNum & "' and rcp09='" & .Fields("rcp09") & "'"
            cnnConnection.Execute strSql, intI
            .MoveNext
         Loop
      End With
   End If
   
   Dim tmpCp27SQL As String

   tmpCp27SQL = "decode(sign(instr(',L,FCL,CFL,LA,LIN,ACS,',','||np02||',')),1,decode(np07,'6001',sqldatet(cp27),'')" & _
      ",decode(sign(instr(',P,PS,CFP,CPS,FCP,FG,',','||np02||',')),1,decode(sign(instr(',997,998,994,995,996,999,411,1204,1503,1209,1603,',','||np07||',')),1,sqldatet(cp27),'')" & _
      ",decode(sign(instr(',994,997,998,995,996,999,305,1403,1701,1711,312,',','||np07||',')),1,sqldatet(cp27),'')))"
   
   Dim tmpKindSql As String
   tmpKindSql = "decode(sign(instr(',L,FCL,CFL,LA,LIN,ACS,',','||np02||',')),1,decode(np07,'6001','4','2')" & _
      ",decode(sign(instr(',P,PS,CFP,CPS,FCP,FG,',','||np02||',')),1,decode(sign(instr(',997,998,994,995,996,999,411,1204,1503,1209,1603,',','||np07||',')),1,'4','2')" & _
      ",decode(sign(instr(',994,997,998,995,996,999,305,1403,1701,1711,312,',','||np07||',')),1,'4','2')))"
   
   'Check1.未處理、未函知 chkNP.未回覆
   'Add By Sindy 2014/6/12
   If Check1.Value = 1 Or chkNP.Value = 1 Then
      strChkSql = ""
'      If Check1.Value = 1 And chkNP.Value = 1 Then
'         strChkSql = ""
'      Else
'         If chkNP.Value = 1 Then '4.未回覆
'            strChkSql = " and " & tmpKindSql & "='4'"
'         Else '2.未處理
'            strChkSql = " and " & tmpKindSql & "='2'"
'         End If
'      End If
   '2014/6/12 END
      
      '2.未處理 4.未回覆
      'Modify By Sindy 2021/8/31 '' as 約定期限 => sqldatet(np23) as 約定期限
      strSql = strMidSql & " select np02,np03,np04,np05,np01,s1.st15 as 管制人部門,cp13," & tmpKindSql & ",'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(np08) as 本所期限,np10 as 管制人, " & _
                  "sqldatet(np09) as 法定期限,sqldatet(np23) as 約定期限,pa47 as 分所號,NVL(NVL(PA05,PA06),PA07) As 案件名稱,np07 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日," & tmpCp27SQL & " as 發文日," & _
                  "pa26 as 申請人,pa09 as 申請國家,pa11 as 申請案號,np22 as 序號,'' as PKey,substr(np15,1,500) as caserem,substr(cmr04,1,500) ctlrem,'' as RECVTYPE " & _
                  "from nextprogress,patent,caseprogress,staff s1,staff s2,t102inform, ctlremark " & _
                  "where np02 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ") and ((" & tmpKindSql & "='2'" & stCon1 & ") or (" & tmpKindSql & "='4'" & Replace(Replace(UCase(stCon1), "NP10", "CP13"), "S1.ST15", "S2.ST15") & ")) " & _
                  "and np06 is null and pa57 is null and np01=cp09(+) and np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+) " & _
                  "and np10=s1.st01(+) and cp13=s2.st01(+) and np01=ti02(+) and np22=ti04(+) and np01=cmr01(+) and np22=cmr02(+) " & _
                  stConST & stConP & strChkSql
      cnnConnection.Execute strSql, intI
      
      'Added by Morgan 2012/8/21
      strSql = strMidSql & " select np02,np03,np04,np05,np01,s1.st15 as 管制人部門,cp13,'2','" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(np08) as 本所期限,np10 as 管制人, " & _
                 "sqldatet(np09) as 法定期限,sqldatet(np23) as 約定期限,pa47 as 分所號,NVL(NVL(PA05,PA06),PA07) As 案件名稱,np07 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日," & tmpCp27SQL & " as 發文日," & _
                 "pa26 as 申請人,pa09 as 申請國家,pa11 as 申請案號,np22 as 序號,'' as PKey,substr(np15,1,500) as caserem,substr(cmr04,1,500) ctlrem,'' as RECVTYPE " & _
                 "from nextprogress,patent,caseprogress,staff s1,staff s2,t102inform, ctlremark " & _
                 "where np02 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ") " & stCon9 & _
                 " and np06 is null and pa57 is null and np01=cp09(+) and np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+) " & _
                 "and np10=s1.st01(+) and cp13=s2.st01(+) and np01=ti02(+) and np22=ti04(+) and np01=cmr01(+) and np22=cmr02(+) " & _
                 stConST & stConP & strChkSql
      cnnConnection.Execute strSql, intI

      strSql = strMidSql & " select np02,np03,np04,np05,np01,s1.st15 as 管制人部門,cp13," & tmpKindSql & ",'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(np08) as 本所期限,np10 as 管制人," & _
                 "sqldatet(np09) as 法定期限,'' as 約定期限,tm34 as 分所號,NVL(NVL(tm05,TM06),TM07) As 案件名稱,np07 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日," & tmpCp27SQL & " as 發文日," & _
                 "tm23 as 申請人,tm10 as 申請國家,tm12 as 申請案號,np22 as 序號,'' as PKey,substr(np15,1,500) as caserem,substr(cmr04,1,500) ctlrem,'' as RECVTYPE " & _
                 "from nextprogress,trademark,caseprogress,staff s1,staff s2,t102inform,ctlremark " & _
                 "where np02 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ") and ((" & tmpKindSql & "='2'" & stCon1 & ") or (" & tmpKindSql & "='4'" & Replace(Replace(UCase(stCon1), "NP10", "CP13"), "S1.ST15", "S2.ST15") & ")) " & _
                 "and np06 is null and tm29 is null and np01=cp09(+) " & _
                 "and np02=tm01(+) and np03=tm02(+) and np04=tm03(+) and np05=tm04(+) and np10=s1.st01(+) and cp13=s2.st01(+) " & _
                 "and np01=ti02(+)  and np22=ti04(+) " & _
                 "and np01=cmr01(+) and np22=cmr02(+) " & stConST & stConT & strChkSql & _
                 "and decode(np02||np07,'T716',tm17,'T102',tm17,'FCT716',tm17,'FCT102',tm17,'TF716',tm17,'TF102',tm17,'Y')='Y' "
      cnnConnection.Execute strSql, intI

      strSql = strMidSql & " select np02,np03,np04,np05,np01,s1.st15 as 管制人部門,cp13," & tmpKindSql & ",'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(np08) as 本所期限,np10 as 管制人," & _
                "sqldatet(np09) as 法定期限,'' as 約定期限,tm34 as 分所號,NVL(NVL(tm05,TM06),TM07) As 案件名稱,np07 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日," & tmpCp27SQL & " as 發文日," & _
                "tm23 as 申請人,tm10 as 申請國家,tm12 as 申請案號,np22 as 序號,'' as PKey,substr(np15,1,500) as caserem,substr(cmr04,1,500) ctlrem,'' as RECVTYPE " & _
                "from nextprogress,trademark,caseprogress,staff s1,staff s2,t102inform,ctlremark " & _
                "where np02 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ") " & stCon1_1 & _
                " and np06 is null and tm29 is null and np01=cp09(+)" & _
                " and np02=tm01(+) and np03=tm02(+) and np04=tm03(+) and np05=tm04(+) and np10=s1.st01(+) and cp13=s2.st01(+)" & _
                " and np01=ti02(+)  and np22=ti04(+) " & _
                " and np01=cmr01(+) and np22=cmr02(+) and np02||to_char(np07) in ('CFT102','CFT105','FCT208') " & stConST & stConT & strChkSql & _
                " and decode(np02||np07,'T716',tm17,'T102',tm17,'FCT716',tm17,'FCT102',tm17,'TF716',tm17,'TF102',tm17,'Y')='Y' "
      cnnConnection.Execute strSql, intI
      
      strSql = strMidSql & " select np02,np03,np04,np05,np01,s1.st15 as 管制人部門,cp13," & tmpKindSql & ",'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(np08) as 本所期限,np10 as 管制人," & _
                 "sqldatet(np09) as 法定期限,'' as 約定期限,sp28 as 分所號,NVL(NVL(SP05,SP06),SP07) As 案件名稱,np07 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日," & tmpCp27SQL & " as 發文日," & _
                 "sp08 as 申請人,sp09 as 申請國家,sp11 as 申請案號,np22 as 序號,'' as PKey,substr(np15,1,500) as caserem,substr(cmr04,1,500) ctlrem,'' as RECVTYPE " & _
                 "from nextprogress,servicepractice,caseprogress,staff s1,staff s2,t102inform,ctlremark " & _
                 "where np02 in (" & SQLGrpStr(GetAllSysKind(systemkind), 5) & ") and ((" & tmpKindSql & "='2'" & stCon1 & ") or (" & tmpKindSql & "='4'" & Replace(Replace(UCase(stCon1), "NP10", "CP13"), "S1.ST15", "S2.ST15") & ")) " & _
                 "and np06 is null and sp15 is null and np01=cp09(+) and np02=sp01(+) and np03=sp02(+) and np04=sp03(+) and np05=sp04(+) and np10=s1.st01(+) and cp13=s2.st01(+) " & _
                 "and np01=ti02(+) and np22=ti04(+) and np01=cmr01(+) and np22=cmr02(+) " & stConST & stConS & strChkSql
      cnnConnection.Execute strSql, intI
      
      
      strSql = strMidSql & " select np02,np03,np04,np05,np01,s1.st15 as 管制人部門,nvl(getlos04(np01),cp13)," & tmpKindSql & ",'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(np08) as 本所期限,np10 as 管制人," & _
                 "sqldatet(np09) as 法定期限,'' as 約定期限,lc16 as 分所號,NVL(NVL(LC05,LC06),LC07) As 案件名稱,np07 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日," & tmpCp27SQL & " as 發文日," & _
                 "lc11 as 申請人,lc15 as 申請國家,'' as 申請案號,np22 as 序號,'' as PKey,substr(np15,1,500) as caserem,substr(cmr04,1,500) ctlrem,'' as RECVTYPE " & _
                 "from nextprogress,lawcase,caseprogress,staff s1,staff s2,t102inform, ctlremark " & _
                 "where np02 in (" & SQLGrpStr(GetAllSysKind(systemkind), 3) & ") " & _
                 "and ((" & tmpKindSql & "='2'" & stCon1 & ") or (" & tmpKindSql & "='4'" & Replace(Replace(UCase(stCon1), "NP10", "CP13"), "S1.ST15", "S2.ST15") & ")) " & _
                 "and np06 is null and lc08 is null and np01=cp09(+) " & _
                 "and np02=lc01(+) and np03=lc02(+) and np04=lc03(+) and np05=lc04(+) " & _
                 "and np10=s1.st01(+) and cp13=s2.st01(+) " & _
                 "and np01=ti02(+) and np22=ti04(+) and np01=cmr01(+) and np22=cmr02(+) " & stConST & stConL & strChkSql
      'Modify By Sindy 2021/7/28 抓案源介紹人 cp13=s2.st01(+) => nvl(getlos04(cp09),cp13)=s2.st01(+)
      strSql = strMidSql & " select np02,np03,np04,np05,np01,s1.st15 as 管制人部門,s2.st01," & tmpKindSql & ",'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(np08) as 本所期限,np10 as 管制人," & _
                 "sqldatet(np09) as 法定期限,'' as 約定期限,lc16 as 分所號,NVL(NVL(LC05,LC06),LC07) As 案件名稱,np07 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日," & tmpCp27SQL & " as 發文日," & _
                 "lc11 as 申請人,lc15 as 申請國家,'' as 申請案號,np22 as 序號,'' as PKey,substr(np15,1,500) as caserem,substr(cmr04,1,500) ctlrem,'' as RECVTYPE " & _
                 "from nextprogress,lawcase,caseprogress,staff s1,staff s2,t102inform, ctlremark " & _
                 "where np02 in (" & SQLGrpStr(GetAllSysKind(systemkind), 3) & ") " & _
                 "and ((" & tmpKindSql & "='2'" & Replace(UCase(stCon1), "NP10", "nvl(getlos04(np01),NP10)") & ") or (" & tmpKindSql & "='4'" & Replace(Replace(UCase(stCon1), "NP10", "nvl(getlos04(np01),cp13)"), "S1.ST15", "S2.ST15") & ")) " & _
                 "and np06 is null and lc08 is null and np01=cp09(+) " & _
                 "and np02=lc01(+) and np03=lc02(+) and np04=lc03(+) and np05=lc04(+) " & _
                 "and nvl(getlos04(np01),np10)=s1.st01(+) and nvl(getlos04(np01),cp13)=s2.st01(+) " & _
                 "and np01=ti02(+) and np22=ti04(+) and np01=cmr01(+) and np22=cmr02(+) " & stConST & stConL & strChkSql
      cnnConnection.Execute strSql, intI
      
      strSql = strMidSql & " select np02,np03,np04,np05,np01,s1.st15 as 管制人部門,nvl(getlos04(np01),cp13)," & tmpKindSql & ",'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(np08) as 本所期限,np10 as 管制人," & _
                "sqldatet(np09) as 法定期限,'' as 約定期限,hc07 as 分所號,HC06 As 案件名稱,np07 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日," & tmpCp27SQL & " as 發文日," & _
                "hc05 as 申請人,'000' as 申請國家,'' as 申請案號,np22 as 序號,'' as PKey,substr(np15,1,500) as caserem,substr(cmr04,1,500) ctlrem,'' as RECVTYPE " & _
                "from nextprogress,hirecase,caseprogress,staff s1,staff s2,t102inform, ctlremark " & _
                "where np02 in (" & SQLGrpStr(GetAllSysKind(systemkind), 4) & ") " & _
                "and ((" & tmpKindSql & "='2'" & stCon1 & ") or (" & tmpKindSql & "='4'" & Replace(Replace(UCase(stCon1), "NP10", "CP13"), "S1.ST15", "S2.ST15") & ")) " & _
                "and np06 is null and hc09 is null and np01=cp09(+) " & _
                "and np02=hc01(+) and np03=hc02(+) and np04=hc03(+) and np05=hc04(+) " & _
                "and np10=s1.st01(+) and cp13=s2.st01(+) " & _
                "and np01=ti02(+) and np22=ti04(+) and np01=cmr01(+) and np22=cmr02(+) " & stConST & stConH & strChkSql
      cnnConnection.Execute strSql, intI
      'Modify By Sindy 2021/7/28 抓案源介紹人 cp13=s2.st01(+) => nvl(getlos04(cp09),cp13)=s2.st01(+)
      '                                       np10 => Replace(UCase(stCon1), "NP10", "nvl(getlos04(np01),NP10)")
      '                                       cp13 => Replace(UCase(stCon1), "NP10", "nvl(getlos04(np01),cp13)")
      strSql = strMidSql & " select np02,np03,np04,np05,np01,s1.st15 as 管制人部門,s2.st01," & tmpKindSql & ",'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(np08) as 本所期限,np10 as 管制人," & _
                "sqldatet(np09) as 法定期限,'' as 約定期限,hc07 as 分所號,HC06 As 案件名稱,np07 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日," & tmpCp27SQL & " as 發文日," & _
                "hc05 as 申請人,'000' as 申請國家,'' as 申請案號,np22 as 序號,'' as PKey,substr(np15,1,500) as caserem,substr(cmr04,1,500) ctlrem,'' as RECVTYPE " & _
                "from nextprogress,hirecase,caseprogress,staff s1,staff s2,t102inform, ctlremark " & _
                "where np02 in (" & SQLGrpStr(GetAllSysKind(systemkind), 4) & ") " & _
                "and ((" & tmpKindSql & "='2'" & Replace(UCase(stCon1), "NP10", "nvl(getlos04(np01),NP10)") & ") or (" & tmpKindSql & "='4'" & Replace(Replace(UCase(stCon1), "NP10", "nvl(getlos04(np01),cp13)"), "S1.ST15", "S2.ST15") & ")) " & _
                "and np06 is null and hc09 is null and np01=cp09(+) " & _
                "and np02=hc01(+) and np03=hc02(+) and np04=hc03(+) and np05=hc04(+) " & _
                "and nvl(getlos04(np01),np10)=s1.st01(+) and nvl(getlos04(np01),cp13)=s2.st01(+) " & _
                "and np01=ti02(+) and np22=ti04(+) and np01=cmr01(+) and np22=cmr02(+) " & stConST & stConH & strChkSql
      cnnConnection.Execute strSql, intI
      
      
      If chkNP.Value = 0 Then '無 4.未回覆
         strSql = "delete from R100123_1 where id='" & strUserNum & "' and rkind='4'"
         cnnConnection.Execute strSql, intI
      ElseIf chkNP.Value = 1 And Check5.Value = 1 Then '有4.未回覆 刪除含此期間法定期限案件(逾本所期限)
         'Modify By Sindy 2021/7/28 會出現錯誤 ORA-01722: 無效的數字 01722. 00000 -  "invalid number"
'         strSql = "delete from R100123_1 where id='" & strUserNum & "' and rkind='4'" & _
'                    " and to_char(to_date(replace(replace(rcp06,'/',''),'*','')+19110000,'YYYYMMDD'),'YYYYMMDD')<" & strSrvDate(1) & " and to_char(to_date(replace(replace(rcp07,'/',''),'*','')+19110000,'YYYYMMDD'),'YYYYMMDD')>=" & DBDATE(txtCloseDate(0)) & " And to_char(to_date(replace(replace(rcp07,'/',''),'*','')+19110000,'YYYYMMDD'),'YYYYMMDD')<=" & DBDATE(txtCloseDate(1))
         strSql = "delete from R100123_1 where rcp01||rcp02||rcp03||rcp04||rcp09||ID||rkind in(" & _
                  "select rcp01||rcp02||rcp03||rcp04||rcp09||ID||rkind from (" & _
                  "SELECT to_char(to_date(replace(replace(rcp06,'/',''),'*','')+19110000,'YYYYMMDD'),'YYYYMMDD') AS Trcp06" & _
                  ",to_char(to_date(replace(replace(rcp07,'/',''),'*','')+19110000,'YYYYMMDD'),'YYYYMMDD') AS Trcp07" & _
                  ",rcp01,rcp02,rcp03,rcp04,rcp09,ID,rkind FROM R100123_1 where id='" & strUserNum & "' and rkind='4'" & _
                  ") Where Trcp06<" & strSrvDate(1) & _
                  " and Trcp07>=" & DBDATE(txtCloseDate(0)) & _
                  " And Trcp07<=" & DBDATE(txtCloseDate(1)) & _
                  ")"
         cnnConnection.Execute strSql, intI
      End If
      If Check1.Value = 0 Then '無 2.未處理
         strSql = "delete from R100123_1 where id='" & strUserNum & "' and rkind='2'"
         cnnConnection.Execute strSql, intI
      End If
   End If
   
   '9.未函知
   If Check1.Value = 1 Then
      strSql = "update R100123_1 set rkind='9' where id='" & strUserNum & "' and rkind='2' and rcp01 in('CFP','P') and rCP10 in ('605','606','607','416','119','930')" & _
                 " and not exists (select cp09 from caseprogress where cp43=rcp09 and cp10='1913' and cp30=RNP22)"
      cnnConnection.Execute strSql, intI

      strSql = "update R100123_1 set rkind='9' where id='" & strUserNum & "' and rkind='2' and rcp01 in('T','TB','TF') and rCP10 in ('105','109','702','708','716')" & _
                 " and not exists (select cp09 from caseprogress where cp43=rcp09 and cp10='1725' and cp30=RNP22)"
      cnnConnection.Execute strSql, intI
      
      'Modified by Lydia 2021/06/04 +VB在O12執行有問題,to_number(replace(replace(RCP06,'/',''),'*',''))=> to_number(ltrim(replace(replace(RCP06,'/',''),'*','')))
      strSql = "update R100123_1 set rkind='9' where id='" & strUserNum & "' and rkind='2' and rcp01 in('T','TF') and rCP10 in ('102')" & _
                 " and not exists (select cp09 from caseprogress where cp43=rcp09 and cp10='1725' and cp30=RNP22)" & _
                 " and to_number(ltrim(replace(replace(RCP06,'/',''),'*','')))>=1050401"
      cnnConnection.Execute strSql, intI
      
      strSql = "update R100123_1 set rkind='9' where id='" & strUserNum & "' and rkind='2' and rcp01 in('CFT') and rCP10 in ('102')" & _
                 " and not exists (select cp09 from caseprogress where cp43=rcp09 and cp10='1717' and cp30=RNP22)"
      cnnConnection.Execute strSql, intI
      
      '1723.本所通知使用宣誓
      '1711.通知使用宣誓
      strSql = "update R100123_1 set rkind='9' where id='" & strUserNum & "' and rkind='2' and rcp01 in('CFT') and rCP10 in ('105')" & _
                 " and not exists (select cp09 from caseprogress where cp43=rcp09 and cp10='1723' and cp30=RNP22)" & _
                 " and not exists (select cp09 from caseprogress where cp09=rcp09 and cp10='1711')"
      cnnConnection.Execute strSql, intI
   End If
   
   'Check3.未通知 Check4.未發文
   If Check3.Value = 1 Or Check4.Value = 1 Then
      If Check3.Value = 1 And Check4.Value = 1 Then
         strChkSql = ""
      Else
         If Check3.Value = 1 Then '5.未通知
            strChkSql = " and substr(cp09,1,1)='C' and cp10<>'9001'"
         Else '1.未發文
            strChkSql = " and (substr(cp09,1,1)<>'C' or cp10='9001')"
         End If
      End If
      
      '1.未發文 5.未通知
      strSql = strMidSql & " select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,decode(substr(cp09,1,1),'C','5','1') as 分類,'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,cp14 as 管制人," & _
                "sqldatet(cp07) as 法定期限,'' as 約定期限,pa47 as 分所號,NVL(NVL(PA05,PA06),PA07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日," & _
                "pa26 as 申請人,pa09 as 申請國家,pa11 as 申請案號,cp66 as 序號,'' as PKey,substr(cp64,1,500) as caserem,substr(cmr04,1,500) ctlrem,decode(substr(cp09,1,1),'B','B類','') as RECVTYPE " & _
                "from patent,caseprogress,staff s2,t102inform, ctlremark " & _
                "where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ") " & stCon2 & strChkSql & " and cp158=0 and cp159=0 " & _
                "and substr(cp09,1,1)<>'D' and pa57 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) " & _
                "and cp13= s2.st01(+) and cp09=ti02(+) and cp09=cmr01(+) and cp66=cmr02(+) " & stConST & stConP
      cnnConnection.Execute strSql, intI

      strSql = strMidSql & " select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,decode(substr(cp09,1,1),'C','5','1') as 分類,'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,cp14 as 管制人," & _
                 "sqldatet(cp07) as 法定期限,'' as 約定期限,tm34 as 分所號,NVL(NVL(tm05,TM06),TM07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日," & _
                 "tm23 as 申請人,tm10 as 申請國家,tm12 as 申請案號,cp66 as 序號,'' as PKey,substr(cp64,1,500) as caserem,substr(cmr04,1,500) ctlrem,decode(substr(cp09,1,1),'B','B類','') as RECVTYPE " & _
                 "from trademark,caseprogress,staff s2,t102inform, ctlremark " & _
                 "where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ") " & stCon2 & strChkSql & " and cp158=0 and cp159=0 " & _
                 "and substr(cp09,1,1)<>'D' and tm29 is null and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & _
                 "and cp13=s2.st01(+) " & _
                 "and cp09=ti02(+) and cp09=cmr01(+) and cp66=cmr02(+) " & stConST & stConT
      cnnConnection.Execute strSql, intI

      strSql = strMidSql & " select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,decode(substr(cp09,1,1),'C','5','1') as 分類,'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,cp14 as 管制人," & _
                "sqldatet(cp07) as 法定期限,'' as 約定期限,tm34 as 分所號,NVL(NVL(tm05,TM06),TM07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日," & _
                "tm23 as 申請人,tm10 as 申請國家,tm12 as 申請案號,cp66 as 序號,'' as PKey,substr(cp64,1,500) as caserem,substr(cmr04,1,500) ctlrem,decode(substr(cp09,1,1),'B','B類','') as RECVTYPE " & _
                "from trademark,caseprogress,staff s2,t102inform, ctlremark " & _
                "where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ") " & stCon2_1 & strChkSql & " and cp158=0 and cp159=0 " & _
                "and substr(cp09,1,1)<>'D' and tm29 is null and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & _
                "and cp13=s2.st01(+) " & _
                "and cp09=ti02(+) and cp01||cp10 in ('CFT102','CFT105','FCT208') and cp09=cmr01(+) and cp66=cmr02(+) " & stConST & stConT
      cnnConnection.Execute strSql, intI

      strSql = strMidSql & " select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,decode(substr(cp09,1,1),'C','5','1') as 分類,'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,cp14 as 管制人," & _
                "sqldatet(cp07) as 法定期限,'' as 約定期限,sp28 as 分所號,NVL(NVL(SP05,SP06),SP07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日," & _
                "sp08 as 申請人,sp09 as 申請國家,sp11 as 申請案號,cp66 as 序號,'' as PKey,substr(cp64,1,500) as caserem,substr(cmr04,1,500) ctlrem,decode(substr(cp09,1,1),'B','B類','') as RECVTYPE " & _
                "from servicepractice,caseprogress,staff s2,t102inform,ctlremark " & _
                "where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 5) & ") " & stCon2 & strChkSql & " and cp158=0 and cp159=0 " & _
                "and substr(cp09,1,1)<>'D' and sp15 is null and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) " & _
                "and cp13=s2.st01(+) and cp09=ti02(+) and cp09=cmr01(+) and cp66=cmr02(+) " & stConST & stConS
      cnnConnection.Execute strSql, intI
      
      
      strSql = strMidSql & " select cp01,cp02,cp03,cp04,cp09,nvl(getlos04(cp09,3),cp12) as 管制人部門,nvl(getlos04(cp09),cp13),decode(cp10,'9001','1',decode(substr(cp09,1,1),'C','5','1')) as 分類,'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,cp14 as 管制人," & _
                "sqldatet(cp07) as 法定期限,'' as 約定期限,lc16 as 分所號,NVL(NVL(LC05,LC06),LC07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日," & _
                "lc11 as 申請人,lc15 as 申請國家,'' as 申請案號,cp66 as 序號,'' as PKey,substr(cp64,1,500) as caserem,substr(cmr04,1,500) ctlrem,decode(substr(cp09,1,1),'B','B類','') as RECVTYPE " & _
                "from lawcase,caseprogress,staff s2,t102inform,ctlremark " & _
                "where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 3) & ") " & stCon2 & strChkSql & " and cp158=0 and cp159=0 " & _
                "and substr(cp09,1,1)<>'D' and lc08 is null and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) " & _
                "and cp13=s2.st01(+) and cp09=ti02(+) and cp09=cmr01(+) and cp66=cmr02(+) " & stConST & stConL
      cnnConnection.Execute strSql, intI
      'Modify By Sindy 2021/7/28 抓案源介紹人 cp13=s2.st01(+) => nvl(getlos04(cp09),cp13)=s2.st01(+)
      strSql = strMidSql & " select cp01,cp02,cp03,cp04,cp09,s2.st15 as 管制人部門,s2.st01,decode(cp10,'9001','1',decode(substr(cp09,1,1),'C','5','1')) as 分類,'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,cp14 as 管制人," & _
                "sqldatet(cp07) as 法定期限,'' as 約定期限,lc16 as 分所號,NVL(NVL(LC05,LC06),LC07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日," & _
                "lc11 as 申請人,lc15 as 申請國家,'' as 申請案號,cp66 as 序號,'' as PKey,substr(cp64,1,500) as caserem,substr(cmr04,1,500) ctlrem,decode(substr(cp09,1,1),'B','B類','') as RECVTYPE " & _
                "from lawcase,caseprogress,staff s2,t102inform,ctlremark " & _
                "where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 3) & ") " & Replace(UCase(stCon2), "CP13", "S2.ST01") & strChkSql & " and cp158=0 and cp159=0 " & _
                "and substr(cp09,1,1)<>'D' and lc08 is null and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) " & _
                "and nvl(getlos04(cp09),cp13)=s2.st01(+) and cp09=ti02(+) and cp09=cmr01(+) and cp66=cmr02(+) " & stConST & stConL
      cnnConnection.Execute strSql, intI
      
      strSql = strMidSql & " select cp01,cp02,cp03,cp04,cp09,nvl(getlos04(cp09,3),cp12) as 管制人部門,nvl(getlos04(cp09),cp13),decode(substr(cp09,1,1),'C','5','1') as 分類,'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,cp14 as 管制人," & _
                 "sqldatet(cp07) as 法定期限,'' as 約定期限,hc07 as 分所號,HC06 As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日," & _
                 "hc05 as 申請人,'000' as 申請國家,'' as 申請案號,cp66 as 序號,'' as PKey,substr(cp64,1,500) as caserem,substr(cmr04,1,500) ctlrem,decode(substr(cp09,1,1),'B','B類','') as RECVTYPE " & _
                "from hirecase,caseprogress,staff s2,t102inform,ctlremark " & _
                "where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 4) & ") " & stCon2 & strChkSql & " and cp158=0 and cp159=0 " & _
                "and substr(cp09,1,1)<>'D' and hc09 is null and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) " & _
                "and cp13=s2.st01(+) and cp09=ti02(+) and cp09=cmr01(+) and cp66=cmr02(+) " & stConST & stConH
      cnnConnection.Execute strSql, intI
      'Modify By Sindy 2021/7/28 抓案源介紹人 cp13=s2.st01(+) => nvl(getlos04(cp09),cp13)=s2.st01(+)
      strSql = strMidSql & " select cp01,cp02,cp03,cp04,cp09,s2.st15 as 管制人部門,s2.st01,decode(substr(cp09,1,1),'C','5','1') as 分類,'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,cp14 as 管制人," & _
                 "sqldatet(cp07) as 法定期限,'' as 約定期限,hc07 as 分所號,HC06 As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日," & _
                 "hc05 as 申請人,'000' as 申請國家,'' as 申請案號,cp66 as 序號,'' as PKey,substr(cp64,1,500) as caserem,substr(cmr04,1,500) ctlrem,decode(substr(cp09,1,1),'B','B類','') as RECVTYPE " & _
                "from hirecase,caseprogress,staff s2,t102inform,ctlremark " & _
                "where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 4) & ") " & Replace(UCase(stCon2), "CP13", "S2.ST01") & strChkSql & " and cp158=0 and cp159=0 " & _
                "and substr(cp09,1,1)<>'D' and hc09 is null and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) " & _
                "and nvl(getlos04(cp09),cp13)=s2.st01(+) and cp09=ti02(+) and cp09=cmr01(+) and cp66=cmr02(+) " & stConST & stConH
      cnnConnection.Execute strSql, intI
   End If
   
   '1.未發文
   If Check4.Value = 1 Then
      'Add By Sindy 2012/6/4 T、FCT台灣商標爭議案逾承辦期限、逾指定會稿日
      'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
      strSql = strMidSql & " select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'1' as 分類,'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,cp14 as 管制人," & _
                 "sqldatet(cp07) as 法定期限,'' as 約定期限,tm34 as 分所號,NVL(NVL(tm05,TM06),TM07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日," & _
                  "tm23 as 申請人,tm10 as 申請國家,tm12 as 申請案號,cp66 as 序號,'' as PKey,substr(cp64,1,500) as caserem, substr(cmr04,1,500) ctlrem,decode(substr(cp09,1,1),'B','B類','') as RECVTYPE " & _
                  "from trademark,caseprogress,staff s2,t102inform,EngineerProgress,ctlremark " & _
                  "where CP05>=20120601 and cp01 in ('T','FCT')  and cp158=0 and cp159=0 and substr(cp09,1,1)<>'D' AND CP10 in (" & TMdebate & ") And Not(cp01='FCT' And InStr(" & FCT_NotTMdebate & ", cp10) > 0) " & _
                  "and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and TM10='000' and tm29 is null " & _
                  " and cp13=s2.st01(+) and cp09=cmr01(+) and cp66=cmr02(+) " & stConST & stCon2 & stConT & " and CP09=EP02(+) and cp09=ti02(+) " & _
                  " and ((CP48<" & strSrvDate(1) & " and CP48 is not null) or (EP28<" & strSrvDate(1) & " and EP28 is not null))"
      cnnConnection.Execute strSql, intI
      'Added by Lydia 2018/12/10 +T台灣案非爭議案
      If strSrvDate(1) >= T案收文齊備啟用日 Then
            'Modified by Lydia 2022/07/15  T大陸案之齊備日管控T大陸案之齊備日管控: tm10='000' => tm10 in ('000','020')
            strSql = strMidSql & " select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'1' as 分類,'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,cp14 as 管制人," & _
                    "sqldatet(cp07) as 法定期限,'' as 約定期限,tm34 as 分所號,NVL(NVL(tm05,TM06),TM07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日," & _
                    "tm23 as 申請人,tm10 as 申請國家,tm12 as 申請案號,cp66 as 序號,'' as PKey,substr(cp64,1,500) as caserem, substr(cmr04,1,500) ctlrem,decode(substr(cp09,1,1),'B','B類','') as RECVTYPE " & _
                    "from trademark,caseprogress,staff s2,t102inform,EngineerProgress,ctlremark " & _
                    "where cp01 ='T' and cp05>=" & T案收文齊備啟用日 & " and cp158=0 and cp159=0 and substr(cp09,1,1)='A' AND CP10 not in (" & TMdebate & ") " & _
                    " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and tm10 in ('000','020') and tm29 is null " & _
                    " and cp13=s2.st01(+) and cp09=cmr01(+) and cp66=cmr02(+) " & stConST & stCon2 & stConT & " and CP09=EP02(+) and cp09=ti02(+) " & _
                    " and ((CP48<" & strSrvDate(1) & " and CP48 is not null) or (EP28<" & strSrvDate(1) & " and EP28 is not null))"
            cnnConnection.Execute strSql, intI
      End If
      'Added by Lydia 2022/07/15 TC案之文件齊備日管控: 臺灣、大陸
            strSql = strMidSql & " select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'1' as 分類,'" & strUserNum & "',decode(ti01,null,' ','*')||sqldatet(cp06) as 本所期限,cp14 as 管制人," & _
                    "sqldatet(cp07) as 法定期限,'' as 約定期限,sp28 as 分所號,NVL(NVL(sp05,sp06),sp07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日," & _
                    "sp08 as 申請人,sp09 as 申請國家,sp11 as 申請案號,cp66 as 序號,'' as PKey,substr(cp64,1,500) as caserem, substr(cmr04,1,500) ctlrem,decode(substr(cp09,1,1),'B','B類','') as RECVTYPE " & _
                    "from servicepractice,caseprogress,staff s2,t102inform,EngineerProgress,ctlremark " & _
                    "where cp01 ='TC' and cp05>=" & T案收文齊備啟用日 & " and cp158=0 and cp159=0 and substr(cp09,1,1)='A' " & _
                    " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and sp09 in ('000','020') and sp15 is null " & _
                    " and cp13=s2.st01(+) and cp09=cmr01(+) and cp66=cmr02(+) " & stConST & stCon2 & stConS & " and CP09=EP02(+) and cp09=ti02(+) " & _
                    " and ((CP48<" & strSrvDate(1) & " and CP48 is not null) or (EP28<" & strSrvDate(1) & " and EP28 is not null))"
            cnnConnection.Execute strSql, intI
     'end 2022/07/15
   End If
   
   '若有輸入查詢條件為系統類別,本所案號,案件性質者,不查詢行事曆資料
   If systemkind <> "ALL" Or (Text1(0) <> "" And Text1(1) <> "") Or txtCP10 <> "" Then
      '不查詢行事曆資料
   Else
      'Modify by Amy 2023/05/09 caserem取250字,避免欄位過長當掉 ex:SS01=A3013 SS02=20230601 SS03=1 的資料
      strSql = strMidSql & " select '' as cp01,'' as cp02,'' as cp03,'' as cp04,'' as cp09,s2.st15 as 管制人部門,ss01,'3' as 分類,'" & strUserNum & "',' '||sqldatet(ss02) as 本所期限,ss01 as 管制人," & _
                " '' as 法定期限,'' as 約定期限,'' as 分所號,substrb(ss04,1,140) As 案件名稱,'' as 案件性質,'' as 承辦人,'' as 收文日,'' as 發文日," & _
                " '' as 申請人,'' as 申請國家,'' as 申請案號,0 as 序號,ss01||'-'||ss02||'-'||ss03 as PKey, SubStr(Replace(SS04,Chr(13)||Chr(10),'＆'),1,250) as caserem,'' as ctlrem,'' as recvtype  " & _
                " from staff_schedule,staff s2 where ss01=s2.st01(+) " & stConST & stCon3
      cnnConnection.Execute strSql, intI
   End If
   
   '6.未收款
   If Check2.Value = 1 Then
      '目前有預定收款日的所有未收款收文資料
'      stVTBw = "select rd01,substrb(max(rd02||(1000+rd03)||rd05),13) rd05" & _
'         " from ReceivablesDay where rd06 is null group by rd01"
      
      '2011/8/25 modify by sonia 因收文即有cp79但開請款單者不會更新cp79故加cp60<'X'
      'Modified by Morgan 2011/11/22 考慮拆收據情形,收據號改用 getunpayno 函數抓未收款收據號
      'Modified by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期(不限制已發文)
      'Modified by Lydia 2021/04/20 調整符合未收款之條件:(國內收據acc0k0)抓CP60 < 'X' + CP79>0(未收金額) 或 (國外請款acc1k0) CP60 < 'X' + Nvl(A1k29,'N')=未結清帳款
      'Modified by Lydia 2025/06/09 a0k32 is null 改用函數判斷：geta0k32type(a0k01)='1'
      strSql = strMidSql & " select cp01,cp02,cp03,cp04,cp09,管制人部門,cp13,分類,mid,本所期限,管制人,法定期限,約定期限,分所號, 案件名稱,案件性質,承辦人, 收文日 , 發文日, 申請人, 申請國家, 申請案號, 序號, pKey, caserem,ctlrem,RECVTYPE" & _
                   " from (select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'6' as 分類,'" & strUserNum & "' as mid,' '||sqldatet(cp06) as 本所期限,'' as 管制人," & _
                   " ' '||sqldatet(cp07) as 法定期限,'' as 約定期限,pa47 as 分所號,NVL(NVL(PA05,PA06),PA07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日," & _
                   "pa26 as 申請人,pa09 as 申請國家,pa11 as 申請案號,cp66 as 序號,'' as PKey,substr(cp64,1,500) as caserem,substr(cmr04,1,500) ctlrem  ,nvl(cu175,2) as cu175,cp60,cp79,decode(substr(cp09,1,1),'B','B類','') as RECVTYPE" & _
                   " from caseprogress,patent,staff s2,customer,ctlremark " & _
                   "where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ") " & stCon8 & " and cp79>0 and cp60 is not null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) " & _
                   "and s2.st01(+)=cp13 and pa09<>'000' and cp09=cmr01(+) and cp66=cmr02(+) " & stConST & stConP & _
                   " and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) ) aa,acc0k0,acc1k0" & _
                   " where cp60=a0k01(+) and cp60=a1k01(+) and ( (Cp60 <'X' And Cp79>0 And Nvl(A0k02,9999999)<(To_Char(Add_Months(Sysdate-1,Cu175 * -1),'YYYYMMDD')-19110000) And geta0k32type(a0k01)='1') Or " & _
                   " (Cp60 >'X' And Nvl(A1k29,'N') <> 'Y' And Nvl(A1k02,9999999)<(To_Char(Add_Months(Sysdate-1,Cu175 * -1),'YYYYMMDD')-19110000)) )"
      cnnConnection.Execute strSql, intI
      'Modified by Lydia 2025/06/09 a0k32 is null 改用函數判斷：geta0k32type(a0k01)='1'
      strSql = strMidSql & " select cp01,cp02,cp03,cp04,cp09,管制人部門,cp13,分類,mid,本所期限,管制人,法定期限,約定期限,分所號, 案件名稱,案件性質,承辦人, 收文日 , 發文日, 申請人, 申請國家, 申請案號, 序號, pKey,caserem,ctlrem,RECVTYPE" & _
                   " from (select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'6' as 分類, '" & strUserNum & "' as mid,' '||sqldatet(cp06) as 本所期限,'' as 管制人," & _
                   " ' '||sqldatet(cp07) as 法定期限,'' as 約定期限,tm34 as 分所號,NVL(NVL(tm05,tm06),tm07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日," & _
                   "tm23 as 申請人,tm10 as 申請國家,tm12 as 申請案號,cp66 as 序號,'' as PKey,substr(cp64,1,500) as caserem, substr(cmr04,1,500) ctlrem ,nvl(cu175,2) as cu175,cp60,cp79,decode(substr(cp09,1,1),'B','B類','') as RECVTYPE" & _
                   " from caseprogress,trademark,staff s2,customer,ctlremark " & _
                   "where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ") " & stCon8 & " and cp79>0 and cp60 is not null and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & _
                   "and s2.st01(+)=cp13 and tm10<>'000' and cp09=cmr01(+) and cp66=cmr02(+) " & stConST & stConT & _
                   " and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) ) aa,acc0k0,acc1k0" & _
                   " where cp60=a0k01(+) and cp60=a1k01(+) and ( (Cp60 <'X' And Cp79>0 And Nvl(A0k02,9999999)<(To_Char(Add_Months(Sysdate-1,Cu175 * -1),'YYYYMMDD')-19110000) And geta0k32type(a0k01)='1') Or " & _
                   " (Cp60 >'X' And Nvl(A1k29,'N') <> 'Y' And Nvl(A1k02,9999999)<(To_Char(Add_Months(Sysdate-1,Cu175 * -1),'YYYYMMDD')-19110000)) )"
      cnnConnection.Execute strSql, intI
      'Modified by Lydia 2025/06/09 a0k32 is null 改用函數判斷：geta0k32type(a0k01)='1'
      strSql = strMidSql & " select cp01,cp02,cp03,cp04,cp09,管制人部門,cp13,分類,mid,本所期限,管制人,法定期限,約定期限,分所號, 案件名稱,案件性質,承辦人, 收文日 , 發文日, 申請人, 申請國家, 申請案號, 序號, pKey,caserem,ctlrem,RECVTYPE" & _
                   " from (select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'6' as 分類,'" & strUserNum & "' as mid,' '||sqldatet(cp06) as 本所期限,'' as 管制人," & _
                   " ' '||sqldatet(cp07) as 法定期限,'' as 約定期限,sp28 as 分所號,NVL(NVL(sp05,sp06),sp07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日," & _
                   "sp08 as 申請人,sp09 as 申請國家,sp11 as 申請案號,cp66 as 序號,'' as PKey,substr(cp64,1,500) as caserem, substr(cmr04,1,500) ctlrem, nvl(cu175,2) as cu175,cp60,cp79,decode(substr(cp09,1,1),'B','B類','') as RECVTYPE" & _
                   " from caseprogress,servicepractice,staff s2,customer,ctlremark " & _
                   "where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 5) & ") " & stCon8 & " and cp79>0 and cp60 is not null " & _
                   "and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and s2.st01(+)=cp13 and sp09<>'000' and cp09=cmr01(+) and cp66=cmr02(+) " & stConST & stConS & _
                   " and substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+) ) aa,acc0k0,acc1k0" & _
                   " where cp60=a0k01(+) and cp60=a1k01(+) and ( (Cp60 <'X' And Cp79>0 And Nvl(A0k02,9999999)<(To_Char(Add_Months(Sysdate-1,Cu175 * -1),'YYYYMMDD')-19110000) And geta0k32type(a0k01)='1') Or " & _
                   " (Cp60 >'X' And Nvl(A1k29,'N') <> 'Y' And Nvl(A1k02,9999999)<(To_Char(Add_Months(Sysdate-1,Cu175 * -1),'YYYYMMDD')-19110000)) )"
      cnnConnection.Execute strSql, intI
      
      'Modified by Lydia 2022/06/13 debug : lc09=>lc15
      'Modified by Lydia 2025/06/09 a0k32 is null 改用函數判斷：geta0k32type(a0k01)='1'
      strSql = strMidSql & " select cp01,cp02,cp03,cp04,cp09,管制人部門,cp13,分類,mid,本所期限,管制人,法定期限,約定期限,分所號, 案件名稱,案件性質,承辦人, 收文日 , 發文日, 申請人, 申請國家, 申請案號, 序號, pKey,caserem,ctlrem,RECVTYPE" & _
                   " from (select cp01,cp02,cp03,cp04,cp09,cp12 as 管制人部門,cp13,'6' as 分類,'" & strUserNum & "' as mid,' '||sqldatet(cp06) as 本所期限,'' as 管制人," & _
                   " ' '||sqldatet(cp07) as 法定期限,'' as 約定期限,lc16 as 分所號,NVL(NVL(lc05,lc06),lc07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日," & _
                   "lc11 as 申請人,lc15 as 申請國家,'' as 申請案號,cp66 as 序號,'' as PKey,substr(cp64,1,500) as caserem,substr(cmr04,1,500) ctlrem, nvl(cu175,2) as cu175,cp60,cp79,decode(substr(cp09,1,1),'B','B類','') as RECVTYPE " & _
                   "from caseprogress,lawcase,staff s2,customer,ctlremark " & _
                   "where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 3) & ") " & stCon8 & " and cp79>0 and cp60 is not null " & _
                   "and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) " & _
                   "and s2.st01(+)=cp13 and lc15<>'000' and cp09=cmr01(+) and cp66=cmr02(+) " & stConST & stConL & _
                   " and substr(lc11,1,8)=cu01(+) and substr(lc11,9,1)=cu02(+) ) aa,acc0k0,acc1k0" & _
                   " where cp60=a0k01(+) and cp60=a1k01(+) and ( (Cp60 <'X' And Cp79>0 And Nvl(A0k02,9999999)<(To_Char(Add_Months(Sysdate-1,Cu175 * -1),'YYYYMMDD')-19110000) And geta0k32type(a0k01)='1') Or " & _
                   " (Cp60 >'X' And Nvl(A1k29,'N') <> 'Y' And Nvl(A1k02,9999999)<(To_Char(Add_Months(Sysdate-1,Cu175 * -1),'YYYYMMDD')-19110000)) )"
      cnnConnection.Execute strSql, intI
      'Modify By Sindy 2021/7/28 抓案源介紹人 cp13=s2.st01(+) => nvl(getlos04(cp09),cp13)=s2.st01(+)
      'Modified by Lydia 2022/06/13 debug : lc09=>lc15
      'Modified by Lydia 2025/06/09 a0k32 is null 改用函數判斷：geta0k32type(a0k01)='1'
      strSql = strMidSql & " select cp01,cp02,cp03,cp04,cp09,管制人部門,cp13,分類,mid,本所期限,管制人,法定期限,約定期限,分所號, 案件名稱,案件性質,承辦人, 收文日 , 發文日, 申請人, 申請國家, 申請案號, 序號, pKey,caserem,ctlrem,RECVTYPE" & _
                   " from (select cp01,cp02,cp03,cp04,cp09,s2.st15 as 管制人部門,s2.st01 as cp13,'6' as 分類,'" & strUserNum & "' as mid,' '||sqldatet(cp06) as 本所期限,'' as 管制人," & _
                   " ' '||sqldatet(cp07) as 法定期限,'' as 約定期限,lc16 as 分所號,NVL(NVL(lc05,lc06),lc07) As 案件名稱,cp10 as 案件性質,cp14 as 承辦人,sqldatet(cp05) as 收文日,sqldatet(cp27) as 發文日," & _
                   "lc11 as 申請人,lc15 as 申請國家,'' as 申請案號,cp66 as 序號,'' as PKey,substr(cp64,1,500) as caserem,substr(cmr04,1,500) ctlrem, nvl(cu175,2) as cu175,cp60,cp79,decode(substr(cp09,1,1),'B','B類','') as RECVTYPE " & _
                   "from caseprogress,lawcase,staff s2,customer,ctlremark " & _
                   "where cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 3) & ") " & Replace(UCase(stCon8), "CP13", "S2.ST01") & _
                   " and cp79>0 and cp60 is not null " & _
                   "and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) " & _
                   "and s2.st01(+)=nvl(getlos04(cp09),cp13) and lc15<>'000' and cp09=cmr01(+) and cp66=cmr02(+) " & stConST & stConL & _
                   " and substr(lc11,1,8)=cu01(+) and substr(lc11,9,1)=cu02(+) ) aa,acc0k0,acc1k0" & _
                   " where cp60=a0k01(+) and cp60=a1k01(+) and ( (Cp60 <'X' And Cp79>0 And Nvl(A0k02,9999999)<(To_Char(Add_Months(Sysdate-1,Cu175 * -1),'YYYYMMDD')-19110000) And geta0k32type(a0k01)='1') Or " & _
                   " (Cp60 >'X' And Nvl(A1k29,'N') <> 'Y' And Nvl(A1k02,9999999)<(To_Char(Add_Months(Sysdate-1,Cu175 * -1),'YYYYMMDD')-19110000)) )"
      cnnConnection.Execute strSql, intI
      'Added by Lydia 2022/06/13 屬於ACS案件中的TIPS及智財報告不列出，等到智權人員收款作業會列出；P.S.目前未收款只抓國外案,以防之後改成含國內案件,還是增加判斷案件檢查
      StrSQLa = "select * from r100123_1 where id= '" & strUserNum & "' and rkind ='6' and rcp01 in ('ACS') "
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         rsA.MoveFirst
         Do While rsA.EOF = False
            If PUB_ChkACSforTIPS(rsA.Fields("rcp01") & rsA.Fields("rcp02") & rsA.Fields("rcp03") & rsA.Fields("rcp04")) = True Then
               strSql = "delete from r100123_1 where id= '" & strUserNum & "' and rkind ='6' and rcp01='" & rsA.Fields("rcp01") & "' and rcp02='" & rsA.Fields("rcp02") & "' and rcp03='" & rsA.Fields("rcp03") & "' and rcp04='" & rsA.Fields("rcp04") & "' and rcp09='" & rsA.Fields("rcp09") & "' "
               cnnConnection.Execute strSql, intI
            End If
            rsA.MoveNext
         Loop
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      'end 2022/06/13
   End If
   
   'Add By Sindy 2014/10/22 未處理:C類NP01->CP09若無CP27時,未處理則以未通知顯示 ex:CFP-26936
   'Modified by Lydia 2019/04/08 nvl(cp27,0)=0 => cp158=0
   strSql = "UPDATE R100123_1" & _
            " set rkind='5'" & _
            " where id='" & strUserNum & "' and rkind='2'" & _
            " and rcp09 in(select rcp09 from R100123_1,caseprogress where id='" & strUserNum & "' and rkind='2'" & _
                            " and rcp09=cp09 and cp158=0 and substr(rcp09,1,1)='C')"
   cnnConnection.Execute strSql, intI
   '2014/10/22 END
   
'Add By Sindy 2015/9/18 10.結案中 : 原為 未處理,未函知 的情形下,可能已填寫結案單, 但仍在結案流程中(程序尚未處理)
   '                                   請將此類資料的事件請改為 結案中
   '                                   結案中, 預設是否勾選與 未處理,未函知 相同
   'Modify By Sindy 2017/7/25 8碼為結案電子表單編號 ==>  and length(np24)=8
   'Modify by Amy 2020/05/18 T/TF延展、續展、第二期註冊費,若已結案且 未勾「結案中」則不顯示
   If Check7.Enabled = True Then
      strSql = "UPDATE R100123_1" & _
               " set rkind='10'" & _
               " where id='" & strUserNum & "' and rkind in('2','9')" & _
               " and exists (select np01,np22,np06,np24 from nextprogress where np01=RCP09 and np22=RNP22" & _
                           " and np06 is null and np24 is not null and length(np24)=8)"
      cnnConnection.Execute strSql, intI
      If Check7.Value = 0 Then
         strSql = "Delete R100123_1" & _
                    " Where id='" & strUserNum & "' and rkind='10'"
         cnnConnection.Execute strSql, intI
         strSql = "Delete R100123_1" & _
                    " Where id='" & strUserNum & "' And RCP01||RCP10 in ('T102','T716','TF102','TF716') And InStr(rcp06,'*')>0"
         cnnConnection.Execute strSql, intI
      End If
   End If

'***********************************
   'Modify By Sindy 2021/9/16 + 類型
   '將「分所號」改到「管制備註」欄後面，原位置改為「類型」欄，
   '當期限來至於案件進度檔且為B類收文號時此欄才顯示「B類」。
   stCon = "select '' as V,decode(s1.st06,'1','北所','2','中所','3','南所','4','高所','其他') as 所別,a0902 as 管制人部門,s3.st02 as 管制人,s1.st02 as 收文智權人員," & _
            "Replace(rcp06,'*','') as 本所期限,rcp07 as 法定期限,decode(rcp01,'','',rcp01||'-'||rcp02||'-'||rcp03||'-'||rcp04) as 本所案號," & _
            "rcasename as 案件名稱,NVL(DECODE(rnation,'000',CPM03,CPM04),rcp10) as 案件性質," & _
            "Decode(Substr(rcp06,1,1),'*','結案中',decode(rkind,'1','未發文','2','未處理','3','行事曆','4','未回覆','5','未通知','6','未收款','7','未續簽','8','未回執','9','未函知','10','結案中','')) as 事件分類," & _
            "rnp23 as 約定期限,recvtype as 類型, rcp64 as 進度備註, rctlrem as 管制備註,rsubno as 分所號, s2.st02 as 承辦人," & _
            "rcp05 as 收文日,rcp27 as 發文日,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as 申請人,nvl(na03,na04) as 申請國家," & _
            "rcaseno as 申請案號,rcp09 as 收文號,rnp22 as 序號,rpkey as PKey" & _
            " from (select distinct * from R100123_1 where id='" & strUserNum & "'),acc090,staff S1,Staff S2,staff S3,nation,customer,casepropertymap" & _
            " where rcp01=cpm01(+) and rcp10=cpm02(+)" & _
            " and rnation=na01(+)" & _
            " and substr(rappid,1,8)=cu01(+) and substr(rappid,9,1)=cu02(+)" & _
            " and rcp12=a0901(+) and rcp13=s1.st01(+) and rcp14=s2.st01(+) and remp=s3.st01(+)" & _
            " order by s1.st06,rcp12,rcp13,rkind,NVL(rnp23,decode(substr(ltrim(rcp06),1,1),'*',substr('0'||substr(ltrim(rcp06),2,length(ltrim(rcp06))),-9,9),substr('0'||ltrim(rcp06),-9,9))),rcp01,rcp02,rcp03,rcp04"
   CheckOC3
   SetDataListWidth
   grdDataList.Rows = 2
   grdDataList.Clear
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open stCon, cnnConnection, adOpenStatic, adLockReadOnly
      LblCntTime.Caption = LblCntTime.Caption & " ~ " & Format(ServerTime, "##:##:##") & " 共 " & .RecordCount & " 筆"
      If .RecordCount > 0 Then
         grdDataList.FixedCols = 0
         Set grdDataList.Recordset = AdoRecordSet3.Clone
         SetDataListWidth
         grdDataList.FixedCols = 10 '固定欄位
         '當天變淺紅
         grdDataList.Visible = False
         For i = 1 To grdDataList.Rows - 1
            grdDataList.row = i
            ' 相關案件性質
            If .Fields("事件分類") = "未回覆" Then
               grdDataList.TextMatrix(i, colCPM) = grdDataList.TextMatrix(i, colCPM) & PUB_GetNextCasePropertyName(grdDataList.TextMatrix(i, colNP01), grdDataList.TextMatrix(i, colNP22), "1")
            End If
            grdDataList.col = colCP06 '本所期限
            '檢查本所案號是否有轉案至他所,若有,則在本所期限前加※符號
            Dim strCP06 As String
            grdDataList.TextMatrix(i, colCP06) = PUB_GetCP10ValueAttachText(grdDataList.TextMatrix(i, colCaseNo), "728", "※", grdDataList.TextMatrix(i, colCP06))
            strCP06 = Trim(grdDataList.Text)
            If Left(strCP06, 1) = "※" Then
               strCP06 = Mid(strCP06, 2)
            End If
            If Left(strCP06, 1) = "*" Then
               strCP06 = Mid(strCP06, 2)
            End If
            Call recovercolor(i) '還原顏色
         Next i
         grdDataList.Visible = True
      Else
         If bolShowMsgBox = True Then
            MsgBox "無符合資料！", vbInformation
         End If
      End If
   End With
   
   doQuery_1 = True
   If bolChgSystemkind = True Then systemkind = strOldSystemkind
   
   If txtSales.Visible = True Then 'Added by Lydia 2021/05/20 排除隱藏
      If txtSales.Enabled = True Then
         txtSales_GotFocus
         txtSales.SetFocus
      End If
   End If  'Added by Lydia 2021/05/20
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical

End Function

'Added by Lydia 2021/05/10 增加管制備註功能：調整版面
Private Sub SetDataListWidth_1()
Dim arrGridHeadWidth
Dim iRow As Integer
   
'A.  固定欄位：所別、部門、管制人、智權人員、本所期限、法定期限、本所案號、案件名稱、案件性質。拉左右捲軸時欄位固定。
                         '(P.S.部門、案件名稱隱藏)。
'B.  非固定欄：事件、約定期限、分所號、備註、管制備註、承辦人、收文日、發文日、申請人、申請國家、申請案號、收文號。
   
   'Modify By Sindy 2021/9/16 + 類型
   '將「分所號」改到「管制備註」欄後面，原位置改為「類型」欄，
   '當期限來至於案件進度檔且為B類收文號時此欄才顯示「B類」。
   arrGridHeadText = Array("V", "所別", "部門", "管制人", "智權人員" _
                      , "本所期限", "法定期限", "本所案號", "案件名稱", "案件性質" _
                     , "事件", "約定期限", "類型", "備註", "管制備註", "分所號" _
                     , "承辦人", "收文日", "發文日", "申請人" _
                     , "申請國家", "申請案號", "收文號", "序號", "PKey")
                     
Dim iDep As String
   
   iDep = PUB_GetST06(strUserNum)
   '所別及分所號都顯示,分所不顯示所別但加入分所號在本所案號後,北所非M51顯示所別不顯示分所號
   If GetStaffDepartment(strUserNum) = "M51" Then
      arrGridHeadWidth = Array(200, 250, 0, 680, 680 _
                               , 900, 850, 1200, 1200, 1000 _
                               , 680, 850, 450, 1700, 1700, 670 _
                               , 800, 850, 850, 1700 _
                               , 800, 1200, 1800, 0, 0)
   Else
      If iDep = "1" Then '北所
        arrGridHeadWidth = Array(200, 250, 0, 680, 680 _
                                 , 900, 850, 1200, 1200, 1000 _
                                 , 680, 850, 450, 1700, 1700, 0 _
                                 , 800, 850, 850, 1700 _
                                 , 800, 1200, 1800, 0, 0)
      Else  '分所
        arrGridHeadWidth = Array(200, 0, 0, 680, 680 _
                                 , 900, 850, 1200, 1200, 1000 _
                                 , 680, 850, 450, 1700, 1700, 670 _
                                 , 800, 850, 850, 1700 _
                                 , 800, 1200, 1800, 0, 0)
      End If
   End If
   
   grdDataList.MergeCells = flexMergeRestrictColumns
   grdDataList.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To grdDataList.Cols - 1
      grdDataList.row = 0
      grdDataList.col = iRow
      grdDataList.Text = arrGridHeadText(iRow)
      grdDataList.ColWidth(iRow) = arrGridHeadWidth(iRow)
      grdDataList.CellAlignment = flexAlignLeftCenter
   Next
   
   '取得Grid的欄位位置
   If colPKey = 0 Then
       colPKey = PUB_MGridGetId("PKey", Me.grdDataList)
       colType = PUB_MGridGetId("事件", Me.grdDataList)
       colCaseNo = PUB_MGridGetId("本所案號", Me.grdDataList)
       colCP06 = PUB_MGridGetId("本所期限", Me.grdDataList)
       colCPM = PUB_MGridGetId("案件性質", Me.grdDataList)
       colNP01 = PUB_MGridGetId("收文號", Me.grdDataList)
       colNP22 = PUB_MGridGetId("序號", Me.grdDataList)
       colCMR04 = PUB_MGridGetId("管制備註", Me.grdDataList) 'Added by Lydia 2021/05/28
   End If
End Sub

'Added by Lydia 2021/05/28 Grid中之管制備註欄內容同步更新
Public Sub UpdateCMR04(ByVal pRow As Integer, ByVal bolUpdate As Boolean, ByVal pCMR04 As String)
'因為可選多筆連續輸入管制備註，存檔後自動帶下一筆；取消則回前畫面，但有勾選但未顯示的資料的勾選符號必須保留，這樣才知道做到哪一筆。
        
    grdDataList.TextMatrix(pRow, 0) = "" '取消勾選V
    If bolUpdate = True Then '是否更新
        grdDataList.TextMatrix(pRow, colCMR04) = Mid(pCMR04, 1, 500)  '限制長度
    End If
    Call recovercolor(pRow)  '還原顏色
End Sub
