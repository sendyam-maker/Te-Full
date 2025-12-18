VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100114_1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人案件查詢"
   ClientHeight    =   5892
   ClientLeft      =   1704
   ClientTop       =   3108
   ClientWidth     =   9300
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5892
   ScaleWidth      =   9300
   Begin VB.CommandButton CmdAP 
      Caption         =   "互惠期間統計"
      Height          =   345
      Index           =   1
      Left            =   6000
      TabIndex        =   56
      Top             =   2424
      Width           =   1395
   End
   Begin VB.CommandButton cmdMemo 
      BackColor       =   &H00C0FFC0&
      Caption         =   "查詢置換字"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   4070
      Style           =   1  '圖片外觀
      TabIndex        =   55
      Top             =   10
      Width           =   1050
   End
   Begin VB.CheckBox Check3 
      Caption         =   "顯示有無案件"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   90
      TabIndex        =   54
      Top             =   210
      Width           =   1665
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "寄發信函-往來記錄"
      Height          =   345
      Index           =   10
      Left            =   7290
      TabIndex        =   51
      Top             =   1560
      Width           =   1845
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件統計"
      Height          =   345
      Index           =   9
      Left            =   3600
      Style           =   1  '圖片外觀
      TabIndex        =   50
      Top             =   420
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CommandButton CmdAP 
      Caption         =   "來訪資料(Word)"
      Height          =   345
      Index           =   0
      Left            =   4500
      TabIndex        =   48
      Top             =   408
      Width           =   1395
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "列印對造資料(&O)"
      Enabled         =   0   'False
      Height          =   300
      Index           =   8
      Left            =   7720
      Style           =   1  '圖片外觀
      TabIndex        =   45
      Top             =   825
      Width           =   1515
   End
   Begin VB.CheckBox Check1 
      Caption         =   "含投資法務開拓資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   7.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5340
      TabIndex        =   43
      Top             =   1440
      Width           =   1770
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   10
      Left            =   6825
      TabIndex        =   8
      Top             =   1140
      Width           =   1600
   End
   Begin VB.OptionButton Option1 
      Caption         =   "E-Mail："
      Height          =   180
      Index           =   3
      Left            =   5910
      TabIndex        =   41
      Top             =   1200
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "聯絡人(&T)"
      Height          =   345
      Index           =   7
      Left            =   5895
      Style           =   1  '圖片外觀
      TabIndex        =   40
      Top             =   420
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "往來記錄(&N)"
      Height          =   345
      Index           =   6
      Left            =   6885
      Style           =   1  '圖片外觀
      TabIndex        =   39
      Top             =   420
      Width           =   1170
   End
   Begin VB.CheckBox ChkPCT 
      Caption         =   "是否顯示PCT 案"
      Height          =   225
      Left            =   3360
      TabIndex        =   11
      Top             =   1770
      Width           =   1635
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "法務進度(&L)"
      Height          =   345
      Index           =   5
      Left            =   8070
      Style           =   1  '圖片外觀
      TabIndex        =   21
      Top             =   405
      Width           =   1170
   End
   Begin VB.CheckBox chk 
      Caption         =   "所有系統類別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   90
      TabIndex        =   38
      Top             =   0
      Width           =   1875
   End
   Begin VB.Frame Frame2 
      Height          =   360
      Left            =   5190
      TabIndex        =   37
      Top             =   735
      Width           =   2475
      Begin VB.OptionButton Option3 
         Caption         =   "字首比對"
         Height          =   180
         Index           =   0
         Left            =   96
         TabIndex        =   5
         Top             =   144
         Width           =   1125
      End
      Begin VB.OptionButton Option3 
         Caption         =   "模糊比對"
         Height          =   180
         Index           =   1
         Left            =   1230
         TabIndex        =   6
         Top             =   144
         Value           =   -1  'True
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
      Height          =   360
      Left            =   2400
      TabIndex        =   36
      Top             =   2370
      Visible         =   0   'False
      Width           =   3195
      Begin VB.OptionButton Option2 
         Caption         =   "日文"
         Height          =   204
         Index           =   2
         Left            =   2265
         TabIndex        =   4
         Top             =   120
         Width           =   732
      End
      Begin VB.OptionButton Option2 
         Caption         =   "英文"
         Height          =   204
         Index           =   1
         Left            =   1185
         TabIndex        =   3
         Top             =   120
         Width           =   732
      End
      Begin VB.OptionButton Option2 
         Caption         =   "中文"
         Height          =   204
         Index           =   0
         Left            =   105
         TabIndex        =   2
         Top             =   120
         Value           =   -1  'True
         Width           =   732
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   2760
      Left            =   30
      TabIndex        =   25
      Top             =   3090
      Width           =   9270
      _ExtentX        =   16341
      _ExtentY        =   4868
      _Version        =   393216
      Cols            =   17
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   17
   End
   Begin VB.OptionButton Option1 
      Caption         =   "代理人國籍："
      Height          =   204
      Index           =   2
      Left            =   -15
      TabIndex        =   17
      Top             =   2775
      Width           =   1395
   End
   Begin VB.OptionButton Option1 
      Caption         =   "代理人名稱："
      Height          =   204
      Index           =   1
      Left            =   72
      TabIndex        =   1
      Top             =   855
      Width           =   1395
   End
   Begin VB.OptionButton Option1 
      Caption         =   "代理人編號："
      Height          =   204
      Index           =   0
      Left            =   72
      TabIndex        =   27
      Top             =   480
      Value           =   -1  'True
      Width           =   1395
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   7
      Left            =   5520
      MaxLength       =   4
      TabIndex        =   15
      Top             =   2085
      Width           =   852
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   9
      Left            =   1425
      MaxLength       =   4
      TabIndex        =   18
      Top             =   2730
      Width           =   1092
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   8
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   16
      Top             =   2415
      Width           =   1092
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   10
      Top             =   1770
      Width           =   492
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   1590
      MaxLength       =   9
      TabIndex        =   0
      Top             =   450
      Width           =   1932
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   6
      Left            =   4320
      MaxLength       =   4
      TabIndex        =   14
      Top             =   2085
      Width           =   852
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   1080
      MaxLength       =   7
      TabIndex        =   12
      Top             =   2085
      Width           =   852
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   1080
      TabIndex        =   9
      Top             =   1440
      Width           =   2772
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   2280
      MaxLength       =   7
      TabIndex        =   13
      Top             =   2085
      Width           =   852
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   3
      Left            =   7416
      Style           =   1  '圖片外觀
      TabIndex        =   26
      Top             =   3735
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找"
      Default         =   -1  'True
      Height          =   345
      Left            =   5140
      Style           =   1  '圖片外觀
      TabIndex        =   19
      Top             =   10
      Width           =   600
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "關係企業"
      Height          =   345
      Index           =   2
      Left            =   7790
      Style           =   1  '圖片外觀
      TabIndex        =   23
      Top             =   10
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件資料"
      Height          =   345
      Index           =   1
      Left            =   6880
      Style           =   1  '圖片外觀
      TabIndex        =   22
      Top             =   10
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "代理人資料"
      Height          =   345
      Index           =   0
      Left            =   5760
      Style           =   1  '圖片外觀
      TabIndex        =   20
      Top             =   10
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   4
      Left            =   8700
      Style           =   1  '圖片外觀
      TabIndex        =   24
      Top             =   10
      Width           =   600
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   330
      Index           =   0
      Left            =   1590
      TabIndex        =   7
      Top             =   795
      Width           =   3555
      VariousPropertyBits=   679493659
      Size            =   "6271;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label12 
      Caption         =   "⊕：不得代理  ▼：無案件"
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Left            =   7530
      TabIndex        =   53
      Top             =   2490
      Width           =   1260
   End
   Begin VB.Label Label11 
      Caption         =   "●：特殊客戶"
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   7536
      TabIndex        =   52
      Top             =   2280
      Width           =   1260
   End
   Begin VB.Label Label10 
      Caption         =   "執行來訪通知資料，請勿開啟Word檔。"
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   2160
      TabIndex        =   49
      Top             =   12
      Width           =   1692
   End
   Begin VB.Label Label1 
      Caption         =   "＊：舊的名稱 ＄：有呆帳"
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   420
      Left            =   7530
      TabIndex        =   47
      Top             =   1920
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "輸入名稱之特取部分, 不要取國家,省份,城市,例：不可輸美商..,廣東..,廣州.."
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   2
      Left            =   90
      TabIndex        =   46
      Top             =   1200
      Width           =   5805
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "註：紅色不可承接案件／黃底為待活化客戶"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   4320
      TabIndex        =   44
      Top             =   2880
      Width           =   3420
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "模糊比對"
      Height          =   180
      Left            =   8490
      TabIndex        =   42
      Top             =   1200
      Width           =   720
   End
   Begin VB.Line Line2 
      X1              =   5280
      X2              =   5400
      Y1              =   2190
      Y2              =   2190
   End
   Begin VB.Line Line1 
      X1              =   2040
      X2              =   2160
      Y1              =   2190
      Y2              =   2190
   End
   Begin VB.Label lbl1 
      Height          =   255
      Index           =   1
      Left            =   2550
      TabIndex        =   35
      Top             =   2760
      Width           =   1700
   End
   Begin VB.Label lbl1 
      Height          =   255
      Index           =   0
      Left            =   2220
      TabIndex        =   34
      Top             =   2445
      Width           =   3735
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Left            =   120
      TabIndex        =   33
      Top             =   2475
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "（1.收文  2.發文）"
      Height          =   180
      Left            =   1680
      TabIndex        =   32
      Top             =   1755
      Width           =   1440
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "查詢別："
      Height          =   180
      Left            =   120
      TabIndex        =   31
      Top             =   1755
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "系統類別：                                                               (ALL：全部)"
      Height          =   180
      Left            =   120
      TabIndex        =   30
      Top             =   1470
      Width           =   4725
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "日期："
      Height          =   180
      Left            =   120
      TabIndex        =   29
      Top             =   2145
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Left            =   3360
      TabIndex        =   28
      Top             =   2145
      Width           =   900
   End
End
Attribute VB_Name = "frm100114_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/05 改成Form2.0 ; GrdDataList改字型=新細明體-ExtB、txt1(1)改成txtFM2(0)
'Memo by Amy 2013/12/04 拿掉查無資料查對造功能(已加查對造) 拿掉中、英、日查詢選項
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
'重整 2005/10/05 nickc
Option Explicit

Dim s As Long, i As Long, j As Long, strSql As String
Dim StrToGrid As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
'Add by Amy 2013/12/02
Dim StrToPrint As String '記錄編號 for 對造列印
Dim strTp(3) As String, ColName() As String
Dim intCounter As Integer, intRecord As Integer, intPage As Integer, kk As Integer, PLeft() As Integer
Dim bolPrint As Boolean '是否有對造
'end 2013/12/04

'Added by Lydia 2018/10/04 來訪通知資料
Dim m_bExec As Boolean '權限
Private Type INVITEM
   IA01 As String  '代理人資料
   IA02 As String  '排序
   IA03 As String  '系統別
   IA04 As String  '系統日之前第5年統計
   IA05 As String  '系統日之前第4年統計
   IA06 As String  '系統日之前第3年統計
   IA07 As String  '系統日之前第2年統計
   IA08 As String  '系統日之前第1年統計
   IA09 As String  '系統日之當年統計
   IA10 As String  '系統日之歷年合計
End Type
Dim m_Item() As INVITEM
Dim iUpper As Integer '陣列的上限
Dim m_YY(0 To 5)  As String  '系統日之前第X年
Dim bolRetry As Boolean '是否已發生Word錯誤且重試
'end 2018/10/04
Dim m_WordLeft As Long, m_WordTop As Long 'Added by Lydia 2019/04/09 Word開啟位置
Dim m_blnColOrderAsc As Boolean 'Add by Amy 2020/09/04 欄位資料由小到大排序
Dim strField() As String 'Add by Amy 2023/03/08
Dim m_pub_QL05 As String 'Add By Sindy 2025/8/13 只記錄於此Form
Public m_strTotKind As String 'Added by Lydia 2025/09/19 選擇統計方式：1-新申請案、2-案件數

'Modify by Amy 2023/08/28 +IsRelation
Private Sub SetDataListWidth(Optional ByVal IsRelation As Boolean = False)
   grdDataList.row = 0
   grdDataList.col = 0: grdDataList.Text = "V"
   grdDataList.ColWidth(0) = 200
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 1: grdDataList.Text = "編號"
   'grdDataList.ColWidth(1) = 1000
   grdDataList.ColWidth(1) = 1200
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 2: grdDataList.Text = "名稱"
   grdDataList.ColWidth(2) = 4000 'Modify by Amy 2014/06/05 改與申請人查詢同 原:4600
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 3: grdDataList.Text = "國籍"
   grdDataList.ColWidth(3) = 1200 'Modify by Amy 2014/06/05 改與申請人查詢同 原:1500
   grdDataList.CellAlignment = flexAlignCenterCenter
   'Modify by Amy 2013/12/04 插入智權人員欄
   grdDataList.col = 4: grdDataList.Text = "智權人員"
   grdDataList.ColWidth(4) = 800
   grdDataList.CellAlignment = flexAlignCenterCenter
   
   grdDataList.col = 5: grdDataList.Text = "狀態"
   grdDataList.ColWidth(5) = 1000
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 6: grdDataList.Text = "備註"
   grdDataList.ColWidth(6) = 2000
   grdDataList.CellAlignment = flexAlignLeftCenter
    'Add by Amy 2013/12/04
   '因查詢服務對造資料需依sp09抓不智權人員資料,故加申請國家
   grdDataList.col = 7: grdDataList.Text = "申請國家"
   grdDataList.ColWidth(7) = 0
   '抓取對造欄位 for 列印
   grdDataList.col = 8: grdDataList.Text = "總收文號"
   grdDataList.ColWidth(8) = 0
   grdDataList.col = 9: grdDataList.Text = "案件性質"
   grdDataList.ColWidth(9) = 0
   grdDataList.col = 10: grdDataList.Text = "收文日"
   grdDataList.ColWidth(10) = 0
   'end 2013/12/04

   'Added by Lydia 2017/02/14 關聯企業
   'Modify by Amy 2019/09/17 改為日期判斷 原:欄位數
   If strSrvDate(1) < 國外部關聯企業啟用日 Then 'Added by Lydia 2017/12/28
        grdDataList.col = 11: grdDataList.Text = "關聯編號"
        grdDataList.ColWidth(11) = 0
        grdDataList.col = 12: grdDataList.Text = "關聯名稱"
        grdDataList.ColWidth(12) = 0
        grdDataList.col = 13: grdDataList.Text = "關聯關係"
        grdDataList.ColWidth(13) = 0
        grdDataList.col = 14: grdDataList.Text = "關聯說明"
        grdDataList.ColWidth(14) = 0
        grdDataList.FixedCols = 0
   End If 'Added by Lydia 2017/12/29
   'end 2017/02/14
   'Modify by Amy 2022/08/19 +ORGN
   grdDataList.col = 15: grdDataList.Text = "ORGN"
   grdDataList.ColWidth(15) = 0
   grdDataList.FixedCols = 0
   'Add by Amy 2019/09/17 +待活化客戶
   grdDataList.col = 16: grdDataList.Text = "待活化客戶"
   grdDataList.ColWidth(16) = 0
   grdDataList.FixedCols = 0
   'end 2022/08/19
   
   'Modify by Amy 2023/08/28 避免沒改到,從strMenu1搬過來
   '關聯企業
   If IsRelation = True Then
      'Modified by Lydia 2017/12/05 改由啟用日控制
      If strSrvDate(1) >= 國外部關聯企業啟用日 Then
         'Added by Lydia 2017/02/14 欄寬調整
         grdDataList.FixedCols = 3 '固定編號和名稱
         Call PUB_SetMSFGridColor(Me.grdDataList, "15") '底色設定為空白
         grdDataList.ColWidth(2) = 1200 '名稱
         grdDataList.ColWidth(3) = 800 '國籍
         grdDataList.ColWidth(6) = 1200 '備註
         grdDataList.ColWidth(11) = 1000 '關聯編號
         grdDataList.ColWidth(12) = 1200 '關聯名稱
         grdDataList.ColWidth(13) = 1200 '關聯關係
         grdDataList.ColWidth(14) = 1200 '關聯說明
         'end 2017/02/14
      End If
   End If
End Sub

'Add by Amy 2023/03/08 變動的欄位
Private Sub GetField()
    ReDim strField(grdDataList.Cols - 1)
    For j = 0 To grdDataList.Cols - 1
        strField(j) = grdDataList.TextMatrix(0, j)
    Next j
End Sub

Private Function GetValue(pFieldN As String) As Integer
    Dim jj As Integer
 
    For jj = 1 To UBound(strField)
        If UCase(strField(jj)) = UCase(pFieldN) Then
            GetValue = jj
            Exit For
        End If
    Next jj
End Function
'end 2023/03/08

Private Sub chk_Click()
   'Add By Cheng 2003/04/01
   '若勾選所有系統類別
   If Me.chk.Value = vbChecked Then
       Me.txt1(2).Text = "ALL"
   '若取消勾選所有系統類別
   Else
       Me.txt1(2).Text = Systemkind_g
   End If
End Sub

'Add by Amy 2023/08/28 整理
Public Sub PubShowNextData()
   Dim strRepCon As String
   
   '列印對造資料
   If cmdState = 8 Then
      strRepCon = txtFM2(0)
      If Option3(0).Value = True Then
         strRepCon = strRepCon & " (字首比對)"
      ElseIf Option3(1).Value = True Then
         strRepCon = strRepCon & " (模糊比對)"
      End If
      cmdOK(cmdState).Enabled = False
   End If
   Call PubShowNextForm(cmdState, Me, grdDataList, strField, _
      IIf(Check3.Value = vbChecked, True, False), IIf(ChkPCT.Value = vbChecked, True, False), _
      txt1(2), txt1(3), txt1(4), txt1(5), txt1(6), txt1(7), txt1(8), txt1(8), , , , strRepCon)
   If cmdState = 8 Then cmdOK(cmdState).Enabled = True
End Sub

'Mark by Amy 2023/08/28 改抓共用
'92.04.16 nick
Public Sub PubShowNextData_Old()
'Add by Amy 2014/05/07
'Dim strTmp As String
'Dim strCaseNo As String '本所案號(for 對造)
'
'   'Modify by Amy 2023/03/08 欄位改動態
'   Select Case cmdState
'      Case 0
'           Me.Enabled = False
'           For i = 1 To grdDataList.Rows - 1
'           grdDataList.col = 0
'           grdDataList.row = i
'           If Trim(grdDataList.Text) = "V" Then
'               grdDataList.col = 0
'               grdDataList.Text = ""
'               'Add By Sindy 2012/3/21
'               grdDataList.col = 1
'               'Add by Amy 2019/09/17 活化客戶顯示整列黃,若有呆帳編號底為紅其他欄為黃
'               If grdDataList.TextMatrix(grdDataList.row, GetValue("待活化客戶")) = "0" And Right(grdDataList.TextMatrix(grdDataList.row, GetValue("編號")), 1) <> "＊" Then
'                    For j = 0 To grdDataList.Cols - 1
'                        '呆帳
'                        If Right(grdDataList.Text, 1) = "$" And j = 1 Then
'                            grdDataList.CellBackColor = &HFF& '紅色
'                        '活化客戶
'                        Else
'                            grdDataList.col = j
'                            grdDataList.CellBackColor = vbYellow
'                        End If
'                    Next
'               'Add by Amy 2019/08/28 +客戶狀態為 遷移不明/解散/廢止/撤銷/停業/死亡 顯示黑底
'               'Modify by Amy 2022/05/23 拿掉 遷移不明 及 停業
'               ElseIf (Left(grdDataList.Text, 1) = "Y" Or Left(grdDataList.Text, 1) = "X" Or Left(grdDataList.Text, 1) = "R") _
'                  And (grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "解散" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "廢止" _
'                   Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "撤銷" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "死亡") Then
'                    For j = 0 To grdDataList.Cols - 1
'                        grdDataList.col = j
'                        grdDataList.CellBackColor = &H0 '黑色
'                        grdDataList.CellForeColor = &HFF00FF '粉紅色
'                    Next j
'               'Modify by Amy 2013/12/10 +判斷對造
'               ElseIf Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "對造" Then
'                  For j = 0 To grdDataList.Cols - 1
'                     grdDataList.col = j
'                     grdDataList.CellBackColor = &H8080FF
'                  Next j
'               Else
'               '2012/3/21 End
'                  For j = 0 To grdDataList.Cols - 1
'                     'add by nickc 2005/12/14
'                     If j <> 1 Then
'                         grdDataList.col = j
'                         grdDataList.CellBackColor = QBColor(15)
'                     End If
'                  Next j
'               End If
'               grdDataList.col = 1
'               If Not IsNull(grdDataList.Text) Then
'                  If fnSaveParentForm(Me) = False Then
'                      Me.Enabled = True
'                      Exit Sub
'                  End If
'                  Screen.MousePointer = vbHourglass
'                  'Modify by Morgan 2007/12/21 加判斷第一碼切不同畫面
'                  'frm100101_10.Show
'                  'frm100101_10.Tag = Pub_RplStr(GrdDataList.Text)
'                  'frm100101_10.StrMenu
'                  strExc(1) = Pub_RplStr(grdDataList.Text)
'                  Select Case Left(strExc(1), 1)
'                     Case "X"
'                        'Add by Morgan 2008/8/11
'                        If Mid(strExc(1), 10, 1) = "-" Then
'                           strExc(1) = Left(strExc(1), 9)
'                        End If
'                        frm100101_11.Show
'                        frm100101_11.Tag = strExc(1)
'                        frm100101_11.StrMenu
'
'                     Case "Y"
'                        'Add by Morgan 2008/8/11
'                        If Mid(strExc(1), 10, 1) = "-" Then
'                           strExc(1) = Left(strExc(1), 9)
'                        End If
'                        frm100101_10.Show
'                        frm100101_10.Tag = strExc(1)
'                        frm100101_10.StrMenu
'
'                     Case "R"
'                        'Modify By Sindy 2009/06/24 判斷是國外或是國內潛在客戶
'                        strExc(0) = "select * from potcustomer where pcu01(+)='" & Left(strExc(1), 8) & "' and pcu02(+)='" & Mid(strExc(1), 9, 1) & "' "
'                        intI = 1
'                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                        strExc(2) = ""
'                        If intI = 1 Then
'                           strExc(2) = "" & RsTemp.Fields(0)
'                        End If
'                        If strExc(2) <> "" Then '國外
'                           frm100101_14.Show
'                           frm100101_14.Tag = strExc(1)
'                           frm100101_14.StrMenu
'                        Else '國內
'                           frm100101_21.Show
'                           frm100101_21.Tag = strExc(1)
'                           frm100101_21.StrMenu
'                        End If
'                     'Add by Amy 2015/03/27 +客戶端平台帳號
'                     Case "平"
'                        'Modify by Amy 2015/04/15 改以平台編號抓權限
'                        If PUB_ChkCustWebLimit(grdDataList.TextMatrix(grdDataList.row, GetValue("收文日")), strUserNum) = True Then
'                           frm100101_27.Show
'                           frm100101_27.Tag = Trim(grdDataList.TextMatrix(grdDataList.row, GetValue("收文日")))
'                           frm100101_27.StrMenu
'                        Else
'                           Me.Show
'                           MsgBox "您無權限查詢此客戶端平台帳號！", vbInformation
'                        End If
'                     'Add By Sindy 2009/07/22
'                     Case Else
'                        'Modify By Sindy 2012/3/21 +不得代理案件之客戶或代理人
'                        If InStr(strExc(1), "-") = 0 Then
'                           frm100101_25.Show
'                           frm100101_25.Tag = strExc(1)
'                           frm100101_25.StrMenu
'                        Else
'                        '2012/3/21 End
'                           frm100101_22.Show
'                           frm100101_22.Tag = strExc(1)
'                           frm100101_22.StrMenu
'                        End If
'                     '2009/07/22 End
'                  End Select
'                  'end 2007/12/21
'                  Screen.MousePointer = vbDefault
'                  Me.Enabled = True
'                  Exit Sub
'               End If
'           End If
'           Next i
'           Me.Enabled = True
'      '案件資料
'      Case 1
'           Me.Enabled = False
'           For i = 1 To grdDataList.Rows - 1
'           grdDataList.col = 0
'           grdDataList.row = i
'           If Trim(grdDataList.Text) = "V" Then
'               grdDataList.col = 0
'               grdDataList.Text = ""
'               'Add By Sindy 2012/3/21
'               grdDataList.col = 1
'               'Add by Amy 2019/09/17 活化客戶顯示整列黃,若有呆帳編號底為紅其他欄為黃
'               If grdDataList.TextMatrix(grdDataList.row, GetValue("待活化客戶")) = "0" And Right(grdDataList.TextMatrix(grdDataList.row, GetValue("編號")), 1) <> "＊" Then
'                    For j = 0 To grdDataList.Cols - 1
'                        '呆帳
'                        If Right(grdDataList.Text, 1) = "$" And j = 1 Then
'                            grdDataList.CellBackColor = &HFF& '紅色
'                        '活化客戶
'                        Else
'                            grdDataList.col = j
'                            grdDataList.CellBackColor = vbYellow
'                        End If
'                    Next
'               'Add by Amy 2019/08/28 +客戶狀態為 遷移不明/解散/廢止/撤銷/停業/死亡 顯示黑底
'               'Modify by Amy 2022/05/23 拿掉 遷移不明 及 停業
'               ElseIf (Left(grdDataList.Text, 1) = "Y" Or Left(grdDataList.Text, 1) = "X" Or Left(grdDataList.Text, 1) = "R") _
'                  And (grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "解散" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "廢止" _
'                  Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "撤銷" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "死亡") Then
'                    For j = 0 To grdDataList.Cols - 1
'                        grdDataList.col = j
'                        grdDataList.CellBackColor = &H0 '黑色
'                        grdDataList.CellForeColor = &HFF00FF '粉紅色
'                    Next j
'               'Modify by Amy 2013/12/10 +判斷對造
'               ElseIf Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "對造" Then
'                  For j = 0 To grdDataList.Cols - 1
'                     grdDataList.col = j
'                     grdDataList.CellBackColor = &H8080FF
'                  Next j
'               Else
'               '2012/3/21 End
'                  For j = 0 To grdDataList.Cols - 1
'                      'add by nickc 2005/12/14
'                      If j <> 1 Then
'                          grdDataList.col = j
'                          grdDataList.CellBackColor = QBColor(15)
'                      End If
'                  Next j
'               End If
'               grdDataList.col = 1
'               If Not IsNull(grdDataList.Text) Then
'                  If fnSaveParentForm(Me) = False Then
'                      Me.Enabled = True
'                      Exit Sub
'                  End If
'
'                  'Modify by Amy 2014/05/07 +以本所案號抓案件資料
'                  If grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "對造" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "其他相關人" Then
'                    strCaseNo = Pub_RplStr(grdDataList.Text)
'                    strTmp = GetPrjPeopleNum1(strCaseNo)
'                  Else
'                    strTmp = Pub_RplStr(grdDataList.Text)
'                  End If
'                  'end 2014/05/07
'
'                  'Modify by Amy 2014/05/07
'                  Select Case UCase(Left(strTmp, 1)) '2014/05/07 原:UCase(Mid(GrdDataList.Text, 1, 1))
'                  Case "X" '申請人
'                     Screen.MousePointer = vbHourglass
'                     With frm100102_2
'                        .Show
'                        .Tag = strTmp '2014/05/07 原:Pub_RplStr(GrdDataList.Text)
'                        'add b nickc 2007/12/21
'                        .ChkPCT = Me.ChkPCT
'
'                        If strCaseNo <> "" Then
'                            .m_CaseNo = strCaseNo
'                            .StrMenu4
'                        Else
'                            'Modify by Morgan 2008/11/27
'                            '為使查詢案件畫面共用條件改參數方式傳遞
'                            '.StrMenu2
'                            .m_Sys = txt1(2)
'                            .m_Type = txt1(3)
'                            .m_Date1 = txt1(4)
'                            .m_Date2 = txt1(5)
'                            .m_Pty1 = txt1(6)
'                            .m_Pty2 = txt1(7)
'                            .m_Cty1 = txt1(8)
'                            .m_Cty2 = txt1(8)
'                            .StrMenu
'                            'end 2008/11/27
'                        End If
'                     End With
'                     'end 2014/05/07
'                     Screen.MousePointer = vbDefault
'
'                  Case "Y" '代理人
'                      Screen.MousePointer = vbHourglass
'                      With frm100114_2
'                        .Show
'                        .Tag = strTmp '2014/05/07 原:Pub_RplStr(GrdDataList.Text)
'                        'add b nickc 2007/12/21
'                        .ChkPCT = Me.ChkPCT
'                        'Modify by Morgan 2008/11/21
'                        '為使查詢案件畫面共用條件改參數方式傳遞
'                        .m_Sys = txt1(2)
'                        .m_Type = txt1(3)
'                        .m_Date1 = txt1(4)
'                        .m_Date2 = txt1(5)
'                        .m_Pty1 = txt1(6)
'                        .m_Pty2 = txt1(7)
'                        .m_Cty1 = txt1(8)
'                        .m_Cty2 = txt1(8)
'                        'end 2008/11/21
'                        .StrMenu
'                      End With
'                      Screen.MousePointer = vbDefault
'                  Case "R"
'                     Me.Show
'                     MsgBox "該編號為潛在客戶不會有案件資料！", vbInformation
'                  Case Else
'                     Me.Show
'                  End Select
'                  Me.Enabled = True
'                  Exit Sub
'               End If
'           End If
'           Next i
'           Me.Enabled = True
'      '關係企業
'      Case 2
'            Me.Enabled = False
'            strExc(9) = "" 'Added by Lydia 2017/08/18 勾選清單
'            'Modified by Lydia 2017/12/05 改由啟用日控制
'            If strSrvDate(1) < 國外部關聯企業啟用日 Then
'                cnnConnection.Execute "delete from r100114 where id='" & strUserNum & "' "
'            End If
'            'end 2017/12/05
'            For i = 1 To grdDataList.Rows - 1
'              grdDataList.col = 0
'              grdDataList.row = i
'              If Trim(grdDataList.Text) = "V" Then
'                  grdDataList.Text = ""
'                  grdDataList.col = 1
'                  Screen.MousePointer = vbHourglass
'                  'Modified by Lydia 2017/12/05 改由啟用日控制
'                  If strSrvDate(1) < 國外部關聯企業啟用日 Then
'                      Call StrMenu(Pub_RplStr(grdDataList.Text))
'                  Else
'                      'Added by Lydia 2017/02/14 抓關聯企業改成模組,暫存R100114_1
'                      'Modified by Lydia 2017/08/18 是否清除先前記錄
'                      'j = PUB_GetR100114_1(Me.Name, Pub_RplStr(GrdDataList.Text))
'                      j = PUB_GetR100114_1(IIf(strExc(9) = "", True, False), Me.Name, Pub_RplStr(grdDataList.Text))
'                      strExc(9) = strExc(9) & IIf(strExc(9) <> "", ",", "") & Pub_RplStr(grdDataList.Text)
'                      'end 2017/08/18
'                  End If
'                  cmdOK(2).Enabled = False
'                  Screen.MousePointer = vbDefault
'              End If
'            Next i
'            'Modified by Lydia 2017/12/05 改由啟用日控制
'            If strSrvDate(1) < 國外部關聯企業啟用日 Then
'               Call StrMenu1
'            Else
'               'Added by Lydia 2017/02/14 抓關聯企業改成模組,暫存R100114_1
'               If j > 1 Then Call StrMenu1
'            End If
'            'end 2017/12/05
'
'            Me.Enabled = True
'      Case 3
'            tmpBol = fnCancelNowFormAndShowParentForm(Me)
'      Case 4
'           fnCloseAllFrm100
'      'add by nickc 2005/10/05 法務進度
'      Case 5
'           Me.Enabled = False
'           For i = 1 To grdDataList.Rows - 1
'           grdDataList.col = 0
'           grdDataList.row = i
'           If Trim(grdDataList.Text) = "V" Then
'               grdDataList.col = 0
'               grdDataList.Text = ""
'               'Add By Sindy 2012/3/21
'               grdDataList.col = 1
'                'Add by Amy 2019/09/17 活化客戶顯示整列黃,若有呆帳編號底為紅其他欄為黃
'               If grdDataList.TextMatrix(grdDataList.row, GetValue("待活化客戶")) = "0" And Right(grdDataList.TextMatrix(grdDataList.row, GetValue("編號")), 1) <> "＊" Then
'                    For j = 0 To grdDataList.Cols - 1
'                        '呆帳
'                        If Right(grdDataList.Text, 1) = "$" And j = 1 Then
'                            grdDataList.CellBackColor = &HFF& '紅色
'                        '活化客戶
'                        Else
'                            grdDataList.col = j
'                            grdDataList.CellBackColor = vbYellow
'                        End If
'                    Next
'               'Add by Amy 2019/08/28 +客戶狀態為 遷移不明/解散/廢止/撤銷/停業/死亡 顯示黑底
'               'Modify by Amy 2022/05/23 拿掉 遷移不明 及 停業
'               ElseIf (Left(grdDataList.Text, 1) = "Y" Or Left(grdDataList.Text, 1) = "X" Or Left(grdDataList.Text, 1) = "R") _
'                  And (grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "解散" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "廢止" _
'                  Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "撤銷" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "死亡") Then
'                    For j = 0 To grdDataList.Cols - 1
'                        grdDataList.col = j
'                        grdDataList.CellBackColor = &H0 '黑色
'                        grdDataList.CellForeColor = &HFF00FF '粉紅色
'                    Next j
'               'Modify by Amy 2013/12/10 +判斷對造
'               ElseIf Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "對造" Then
'                  For j = 0 To grdDataList.Cols - 1
'                     grdDataList.col = j
'                     grdDataList.CellBackColor = &H8080FF
'                  Next j
'               Else
'               '2012/3/21 End
'                  For j = 0 To grdDataList.Cols - 1
'                      'add by nickc 2005/12/14
'                      If j <> 1 Then
'                          grdDataList.col = j
'                          grdDataList.CellBackColor = QBColor(15)
'                      End If
'                  Next j
'               End If
'               grdDataList.col = 1
'               If Not IsNull(grdDataList.Text) Then
'                  If fnSaveParentForm(Me) = False Then
'                      Me.Enabled = True
'                      Exit Sub
'                  End If
'                  If UCase(Mid(grdDataList.Text, 1, 1)) = "X" Then
'                  '申請人
'                     Screen.MousePointer = vbHourglass
'                     With frm100102_2
'                     .Show
'                     .Tag = Pub_RplStr(grdDataList.Text)
'                     'add b nickc 2007/12/21
'                     .ChkPCT = Me.ChkPCT
'                     .bolIsL = True
'                     'Modify by Morgan 2008/11/27
'                     '為使查詢案件畫面共用條件改參數方式傳遞
'                     '.StrMenu2
'                     .m_Sys = txt1(2)
'                     .m_Type = txt1(3)
'                     .m_Date1 = txt1(4)
'                     .m_Date2 = txt1(5)
'                     .m_Pty1 = txt1(6)
'                     .m_Pty2 = txt1(7)
'                     .m_Cty1 = txt1(8)
'                     .m_Cty2 = txt1(8)
'                     .StrMenu
'                     'end 2008/11/27
'                     End With
'                     Screen.MousePointer = vbDefault
'
'                  Else
'                  '代理人
'                      Screen.MousePointer = vbHourglass
'                      With frm100114_2
'                      .Show
'                      .Tag = Pub_RplStr(grdDataList.Text)
'                      'add b nickc 2007/12/21
'                      .ChkPCT = Me.ChkPCT
'                      .bolIsL = True
'                      'Add by Morgan 2008/11/21
'                      .m_Sys = txt1(2)
'                      .m_Type = txt1(3)
'                      .m_Date1 = txt1(4)
'                      .m_Date2 = txt1(5)
'                      .m_Pty1 = txt1(6)
'                      .m_Pty2 = txt1(7)
'                      .m_Cty1 = txt1(8)
'                      .m_Cty2 = txt1(8)
'                      'end 2008/11/21
'                      .StrMenu
'                      End With
'                      Screen.MousePointer = vbDefault
'                  End If
'                  Me.Enabled = True
'                  Exit Sub
'               End If
'           End If
'           Next i
'           Me.Enabled = True
'      'Add by Morgan 2007/12/18
'      Case 6 '往來記錄
'            Me.Enabled = False
'            For i = 1 To grdDataList.Rows - 1
'            grdDataList.col = 0
'            grdDataList.row = i
'            If Trim(grdDataList.Text) = "V" Then
'               grdDataList.col = 0
'               grdDataList.Text = ""
'               'Add By Sindy 2012/3/21
'               grdDataList.col = 1
'                'Add by Amy 2019/09/17 活化客戶顯示整列黃,若有呆帳編號底為紅其他欄為黃
'               If grdDataList.TextMatrix(grdDataList.row, GetValue("待活化客戶")) = "0" And Right(grdDataList.TextMatrix(grdDataList.row, GetValue("編號")), 1) <> "＊" Then
'                    For j = 0 To grdDataList.Cols - 1
'                        '呆帳
'                        If Right(grdDataList.Text, 1) = "$" And j = 1 Then
'                            grdDataList.CellBackColor = &HFF& '紅色
'                        '活化客戶
'                        Else
'                            grdDataList.col = j
'                            grdDataList.CellBackColor = vbYellow
'                        End If
'                    Next
'               'Add by Amy 2019/08/28 +客戶狀態為 遷移不明/解散/廢止/撤銷/停業/死亡 顯示黑底
'               'Modify by Amy 2022/05/23 拿掉 遷移不明 及 停業
'               ElseIf (Left(grdDataList.Text, 1) = "Y" Or Left(grdDataList.Text, 1) = "X" Or Left(grdDataList.Text, 1) = "R") _
'                  And (grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "解散" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "廢止" _
'                  Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "撤銷" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "死亡") Then
'                    For j = 0 To grdDataList.Cols - 1
'                        grdDataList.col = j
'                        grdDataList.CellBackColor = &H0 '黑色
'                        grdDataList.CellForeColor = &HFF00FF '粉紅色
'                    Next j
'               'Modify by Amy 2013/12/10 +判斷對造
'               ElseIf Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "對造" Then
'                  For j = 0 To grdDataList.Cols - 1
'                     grdDataList.col = j
'                     grdDataList.CellBackColor = &H8080FF
'                  Next j
'               Else
'               '2012/3/21 End
'                  For j = 0 To grdDataList.Cols - 1
'                     If j <> 1 Then
'                         grdDataList.col = j
'                         grdDataList.CellBackColor = QBColor(15)
'                     End If
'                  Next j
'               End If
'               If fnSaveParentForm(Me) = False Then
'                   Me.Enabled = True
'                   Exit Sub
'               End If
'               grdDataList.col = 1
'               Screen.MousePointer = vbHourglass
'               strExc(1) = Pub_RplStr(grdDataList.Text)
'
'               'Modify By Sindy 2010/02/23 判斷是國外或是國內潛在客戶
'               '客戶檔
'               strExc(3) = "select cu12,cu13 from customer where cu01(+)='" & Left(strExc(1), 8) & "' and cu02(+)='" & Mid(strExc(1), 9, 1) & "' "
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(3))
'               strExc(4) = ""
'               If intI = 1 Then
'                  strExc(4) = "" & RsTemp.Fields("cu12")
'               End If
'               '潛在客戶檔
'               strExc(0) = "select * from potcustomer where pcu01(+)='" & Left(strExc(1), 8) & "' and pcu02(+)='" & Mid(strExc(1), 9, 1) & "' "
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               strExc(2) = ""
'               If intI = 1 Then
'                  strExc(2) = "" & RsTemp.Fields(0)
'               End If
''               If strExc(2) <> "" Or Left(Trim(strExc(1)), 1) = "Y" Or Left(Trim(strExc(4)), 1) = "F" Then '國外
'                  frm100101_15.Show
'                  frm100101_15.Tag = strExc(1)
'                  'Modify By Sindy 2020/5/19
'                  'Modify By Sindy 2021/3/25 + Or Left(Trim(strTmp), 1) = "平"
'                  'Modify by Amy 2023/03/09 改與申請人查詢一致
'                  'If strExc(2) <> "" Or Left(Trim(strExc(1)), 1) = "Y" Or Left(Trim(strExc(4)), 1) = "F" Or Left(Trim(strTmp), 1) = "平" Then '國外
'                  'modify by sonia 2023/3/16 有國外代理人權限即可查國外代理人往來記錄,黃美珍77027查Y20894會跑到國內StrMenu2檢查權限
'                  'If strExc(2) <> "" Or _
'                     (Left(Trim(strExc(1)), 1) = "Y" And Left(Pub_StrUserSt03, 1) = "F") Or _
'                     Left(Trim(strExc(4)), 1) = "F" Or Pub_StrUserSt03 = "M51" Or Left(Trim(strTmp), 1) = "平" Then '國外
'                  If strExc(2) <> "" Or _
'                     (Left(Trim(strExc(1)), 1) = "Y" And Left(Pub_StrUserSt03, 1) = "F") Or _
'                     (Left(Trim(strExc(1)), 1) = "Y" And CheckUse("frm100114_1", strExec)) Or _
'                     Left(Trim(strExc(4)), 1) = "F" Or Pub_StrUserSt03 = "M51" Or Left(Trim(strTmp), 1) = "平" Then '國外
'                     frm100101_15.m_quyDataKind = 0 '國外
'                     frm100101_15.StrMenu
'                  Else
'                     frm100101_15.m_quyDataKind = 1 '國內
'                     frm100101_15.StrMenu2
'                  End If
'                  '2020/5/19 END
''               Else '國內
''                  frm100101_20.Show
''                  frm100101_20.Tag = strExc(1)
''                  frm100101_20.StrMenu
''               End If
'
'               Screen.MousePointer = vbDefault
'               grdDataList.col = 0
'               grdDataList.Text = ""
'               'Add By Sindy 2012/3/21
'               grdDataList.col = 1
'               'Add by Amy 2019/09/17 活化客戶顯示整列黃,若有呆帳編號底為紅其他欄為黃
'               If grdDataList.TextMatrix(grdDataList.row, GetValue("待活化客戶")) = "0" And Right(grdDataList.TextMatrix(grdDataList.row, GetValue("編號")), 1) <> "＊" Then
'                    For j = 0 To grdDataList.Cols - 1
'                        '呆帳
'                        If Right(grdDataList.Text, 1) = "$" And j = 1 Then
'                            grdDataList.CellBackColor = &HFF& '紅色
'                        '活化客戶
'                        Else
'                            grdDataList.col = j
'                            grdDataList.CellBackColor = vbYellow
'                        End If
'                    Next
'               'Add by Amy 2019/08/28 +客戶狀態為 遷移不明/解散/廢止/撤銷/停業/死亡 顯示黑底
'               'Modify by Amy 2022/05/23 拿掉 遷移不明 及 停業
'               ElseIf (Left(grdDataList.Text, 1) = "Y" Or Left(grdDataList.Text, 1) = "X" Or Left(grdDataList.Text, 1) = "R") _
'                  And (grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "解散" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "廢止" _
'                  Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "撤銷" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "死亡") Then
'                    For j = 0 To grdDataList.Cols - 1
'                        grdDataList.col = j
'                        grdDataList.CellBackColor = &H0 '黑色
'                        grdDataList.CellForeColor = &HFF00FF '粉紅色
'                    Next j
'               'Modify by Amy 2013/12/10 +判斷對造
'               ElseIf Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "對造" Then
'                  For j = 0 To grdDataList.Cols - 1
'                     grdDataList.col = j
'                     grdDataList.CellBackColor = &H8080FF
'                  Next j
'               Else
'               '2012/3/21 End
'                  For j = 0 To grdDataList.Cols - 1
'                     If j <> 1 Then
'                         grdDataList.col = j
'                         grdDataList.CellBackColor = QBColor(15)
'                     End If
'                  Next j
'               End If
'               Me.Enabled = True
'               Exit Sub
'            End If
'            Next i
'            Me.Enabled = True
'      'Add by Morgan 2008/7/23
'      Case 7 '聯絡人
'            Me.Enabled = False
'            For i = 1 To grdDataList.Rows - 1
'            grdDataList.col = 0
'            grdDataList.row = i
'            If Trim(grdDataList.Text) = "V" Then
'               grdDataList.col = 0
'               grdDataList.Text = ""
'               'Add By Sindy 2012/3/21
'               grdDataList.col = 1
'               'Add by Amy 2019/09/17 活化客戶顯示整列黃,若有呆帳編號底為紅其他欄為黃
'               If grdDataList.TextMatrix(grdDataList.row, GetValue("待活化客戶")) = "0" And Right(grdDataList.TextMatrix(grdDataList.row, GetValue("編號")), 1) <> "＊" Then
'                    For j = 0 To grdDataList.Cols - 1
'                        '呆帳
'                        If Right(grdDataList.Text, 1) = "$" And j = 1 Then
'                            grdDataList.CellBackColor = &HFF& '紅色
'                        '活化客戶
'                        Else
'                            grdDataList.col = j
'                            grdDataList.CellBackColor = vbYellow
'                        End If
'                    Next
'               'Add by Amy 2019/08/28 +客戶狀態為 遷移不明/解散/廢止/撤銷/停業/死亡 顯示黑底
'               'Modify by Amy 2022/05/23 拿掉 遷移不明 及 停業
'               ElseIf (Left(grdDataList.Text, 1) = "Y" Or Left(grdDataList.Text, 1) = "X" Or Left(grdDataList.Text, 1) = "R") _
'                  And (grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "解散" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "廢止" _
'                  Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "撤銷" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "死亡") Then
'                    For j = 0 To grdDataList.Cols - 1
'                        grdDataList.col = j
'                        grdDataList.CellBackColor = &H0 '黑色
'                        grdDataList.CellForeColor = &HFF00FF '粉紅色
'                    Next j
'               'Modify by Amy 2013/12/10 +判斷對造
'               ElseIf Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "對造" Then
'                  For j = 0 To grdDataList.Cols - 1
'                     grdDataList.col = j
'                     grdDataList.CellBackColor = &H8080FF
'                  Next j
'               Else
'               '2012/3/21 End
'                  For j = 0 To grdDataList.Cols - 1
'                     If j <> 1 Then
'                         grdDataList.col = j
'                         grdDataList.CellBackColor = QBColor(15)
'                     End If
'                  Next j
'               End If
'               If fnSaveParentForm(Me) = False Then
'                   Me.Enabled = True
'                   Exit Sub
'               End If
'               grdDataList.col = 1
'               Screen.MousePointer = vbHourglass
'               strExc(1) = Pub_RplStr(grdDataList.Text)
'               'Modify by Morgan 2008/8/5 國內外客戶跑不同畫面
'               Select Case Left(strExc(1), 1)
'                  'Add by Morgan 2008/9/1 潛在客戶跑申請人資料畫面
'                  Case "R"
'                     frm100101_14.Show
'                     frm100101_14.Tag = strExc(1)
'                     frm100101_14.StrMenu
'
'                  Case Else
'                     strExc(2) = "F"
'                     If Left(strExc(1), 1) = "X" Then
'                        strExc(0) = "select st03 from customer,staff where cu01(+)='" & Left(strExc(1), 8) & "' and cu02(+)='" & Mid(strExc(1), 9, 1) & "' and st01(+)=cu13"
'                        intI = 1
'                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                        If intI = 1 Then
'                           strExc(2) = "" & RsTemp.Fields(0)
'                        End If
'                     End If
'                     If Left(strExc(2), 1) = "F" Then
'                        frm100101_17.Show
'                        frm100101_17.Tag = strExc(1)
'                        frm100101_17.StrMenu
'                     Else
'                        frm100101_18.Show
'                        frm100101_18.Tag = strExc(1)
'                        frm100101_18.StrMenu
'                     End If
'               End Select
'               'end 2008/8/5
'
'               Screen.MousePointer = vbDefault
'               grdDataList.col = 0
'               grdDataList.Text = ""
'               'Add By Sindy 2012/3/21
'               grdDataList.col = 1
'               'Add by Amy 2019/09/17 活化客戶顯示整列黃,若有呆帳編號底為紅其他欄為黃
'               If grdDataList.TextMatrix(grdDataList.row, GetValue("待活化客戶")) = "0" And Right(grdDataList.TextMatrix(grdDataList.row, GetValue("編號")), 1) <> "＊" Then
'                    For j = 0 To grdDataList.Cols - 1
'                        '呆帳
'                        If Right(grdDataList.Text, 1) = "$" And j = 1 Then
'                            grdDataList.CellBackColor = &HFF& '紅色
'                        '活化客戶
'                        Else
'                            grdDataList.col = j
'                            grdDataList.CellBackColor = vbYellow
'                        End If
'                    Next
'               'Add by Amy 2019/08/28 +客戶狀態為 遷移不明/解散/廢止/撤銷/停業/死亡 顯示黑底
'               'Modify by Amy 2022/05/23 拿掉 遷移不明 及 停業
'               ElseIf (Left(grdDataList.Text, 1) = "Y" Or Left(grdDataList.Text, 1) = "X" Or Left(grdDataList.Text, 1) = "R") _
'                  And (grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "解散" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "廢止" _
'                  Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "撤銷" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "死亡") Then
'                    For j = 0 To grdDataList.Cols - 1
'                        grdDataList.col = j
'                        grdDataList.CellBackColor = &H0 '黑色
'                        grdDataList.CellForeColor = &HFF00FF '粉紅色
'                    Next j
'               'Modify by Amy 2013/12/10 +判斷對造
'               ElseIf Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "對造" Then
'                  For j = 0 To grdDataList.Cols - 1
'                     grdDataList.col = j
'                     grdDataList.CellBackColor = &H8080FF
'                  Next j
'               Else
'               '2012/3/21 End
'                  For j = 0 To grdDataList.Cols - 1
'                     If j <> 1 Then
'                         grdDataList.col = j
'                         grdDataList.CellBackColor = QBColor(15)
'                     End If
'                  Next j
'               End If
'               Me.Enabled = True
'               Exit Sub
'            End If
'            Next i
'            Me.Enabled = True
'      'Add by Amy 2013/12/04
'      Case 8 '列印對造資料
'            'Modify by Amy 2014/02/25 改印暫存資料
'            'PrintDataA4
'            PrintDataA4_Temp
'            'end 2014/02/25
'      'Move by Lydia 2018/11/09 原本為獨立按鈕,改加入OK
'      Case 9 '案件統計
'            'add by sonia 2018/10/31
'            'Private Sub cmdCase_Click() 'Remove by Lydia 2018/11/09
'               Me.Enabled = False
'               For i = 1 To grdDataList.Rows - 1
'               grdDataList.col = 0
'               grdDataList.row = i
'               If Trim(grdDataList.Text) = "V" Then
'                  grdDataList.col = 0
'                  grdDataList.Text = ""
'                  grdDataList.col = 1
'                  'Add by Amy 2019/08/28 目前程式未完成,導致按了無法回原畫面,Grd變色程式改至最後做
'                  If Left(Pub_RplStr(grdDataList.Text), 1) = "X" Then
'                    Me.Enabled = True
'                    Exit Sub
'                  End If
'                  'end 2019/08/28
''                  For j = 0 To GrdDataList.Cols - 1
''                     If j <> 1 Then
''                         GrdDataList.col = j
''                         GrdDataList.CellBackColor = QBColor(15)
''                     End If
''                  Next j
'                  'end 2019/08/28
'                  grdDataList.col = 1
'                  If Not IsNull(grdDataList.Text) Then
'                     If fnSaveParentForm(Me) = False Then
'                        Me.Enabled = True
'                        Exit Sub
'                     End If
'                     Screen.MousePointer = vbHourglass
'                     strExc(1) = Pub_RplStr(grdDataList.Text)
'                     Select Case Left(strExc(1), 1)
'                        Case "X"
'                           If Mid(strExc(1), 10, 1) = "-" Then
'                              strExc(1) = Left(strExc(1), 9)
'                           End If
'
'            '               frm100114_6.Show
'            '               frm100114_6.Tag = strExc(1)
'            '               frm100114_6.StrMenu
'
'                        Case "Y"
'                           If Mid(strExc(1), 10, 1) = "-" Then
'                              strExc(1) = Left(strExc(1), 9)
'                           End If
'                           frm100114_6.Show
'                           frm100114_6.Tag = strExc(1)
'                           frm100114_6.StrMenu
'
'                        Case Else
'                           Me.Show
'                           grdDataList.col = 0
'                           grdDataList.Text = "V"
'                           MsgBox "非客戶或代理人，無案件統計功能！", vbInformation
'                     End Select
'                     Screen.MousePointer = vbDefault
'                     Me.Enabled = True
'                     Exit Sub
'                  End If
'                  'Add by Amy 2019/08/28
'                  'Add by Amy 2019/09/17 活化客戶顯示整列黃,若有呆帳編號底為紅其他欄為黃
'                  If grdDataList.TextMatrix(grdDataList.row, GetValue("待活化客戶")) = "0" And Right(grdDataList.TextMatrix(grdDataList.row, GetValue("編號")), 1) <> "＊" Then
'                        For j = 0 To grdDataList.Cols - 1
'                            '呆帳
'                            If Right(grdDataList.Text, 1) = "$" And j = 1 Then
'                                grdDataList.CellBackColor = &HFF& '紅色
'                            '活化客戶
'                            Else
'                                grdDataList.col = j
'                                grdDataList.CellBackColor = vbYellow
'                            End If
'                        Next
'                  '客戶狀態為 遷移不明/解散/廢止/撤銷/停業/死亡 顯示黑底
'                  'Modify by Amy 2022/05/23 拿掉 遷移不明 及 停業
'                  ElseIf (Left(grdDataList.Text, 1) = "Y" Or Left(grdDataList.Text, 1) = "X" Or Left(grdDataList.Text, 1) = "R") _
'                    And (grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "解散" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "廢止" _
'                    Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "撤銷" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "死亡") Then
'                      For j = 0 To grdDataList.Cols - 1
'                          grdDataList.col = j
'                          grdDataList.CellBackColor = &H0 '黑色
'                          grdDataList.CellForeColor = &HFF00FF '粉紅色
'                      Next j
'                  '判斷對造
'                  ElseIf Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "對造" Then
'                    For j = 0 To grdDataList.Cols - 1
'                       grdDataList.col = j
'                       grdDataList.CellBackColor = &H8080FF
'                    Next j
'                  Else
'                    For j = 0 To grdDataList.Cols - 1
'                       If j <> 1 Then
'                           grdDataList.col = j
'                           grdDataList.CellBackColor = QBColor(15)
'                       End If
'                    Next j
'                  End If
'                  'end 2019/08/28
'               End If
'               Next i
'               Me.Enabled = True
'            'End Sub 'Remove by Lydia 2018/11/09
'            'end 2018/10/31
'      'Add By Sindy 2019/10/8
'      Case 10 '寄發信函-往來記錄
'         Me.Enabled = False
'         For i = 1 To grdDataList.Rows - 1
'           grdDataList.col = 0
'           grdDataList.row = i
'           If Trim(grdDataList.Text) = "V" Then
'               Screen.MousePointer = vbHourglass
'               grdDataList.Text = ""
'               grdDataList.col = 1
'               strTmp = Trim(grdDataList.TextMatrix(i, GetValue("編號")))
'               If Len(strTmp) = 9 Or (Len(strTmp) = 12 And InStr(strTmp, "-") > 0) Then
'                  Me.Hide
'                  Set frm880022.m_PrevF = Me
'                  frm880022.m_strNo = Left(strTmp, 9)
'                  frm880022.m_PCC02 = IIf(InStr(strTmp, "-") > 0, Right(strTmp, 2), "")
'                  If frm880022.QueryData = True Then
'                     frm880022.Show 'vbModal
'                  End If
'                  Screen.MousePointer = vbDefault
'                  Me.Enabled = True
'                  Exit Sub
'               End If
'           End If
'         Next i
'         Screen.MousePointer = vbDefault
'         Me.Enabled = True
'      '2019/10/8 END
'      Case Else
'   End Select
End Sub

Private Sub cmdMemo_Click()
   cmdState = 99
   If fnSaveParentForm(Me) = False Then
      Me.Enabled = True
      Exit Sub
   End If
   Me.Enabled = False
   Set frm100137.UpForm = Me
   frm100137.Show
   Me.Enabled = True
End Sub

Private Sub cmdok_Click(Index As Integer)
   'add by nickc 2007/01/12
   If Len(Trim(Me.txt1(2).Text)) = 0 Then
       Me.txt1(2).Text = "ALL"
   End If
   
   'Added by Lydia 2025/09/19
   If Index = 9 Then
JumpToReInput1:
      'Modified by Lydia 2025/11/11 改選項說明:1-新申請案    2-案件數=>1-新案（委任申請案）2-在案（目前代理案）
      'm_strTotKind = InputBox("請輸入統計方式：1-新申請案    2-案件數" & vbCrLf & "空白=取消", "案件統計", "1")
      m_strTotKind = InputBox("請輸入統計方式：1-新案（委任申請案）" & vbCrLf & "  2-在案（目前代理案）　　空白=取消", "案件統計", "1")
      If m_strTotKind = "" Then
         Exit Sub
      Else
         If m_strTotKind <> "1" And m_strTotKind <> "2" Then
            GoTo JumpToReInput1
         End If
      End If
   End If
   'end 2025/09/19
   
   '92.04.16 nick 紀錄作用按鍵
   cmdState = Index
   PubShowNextData
   Exit Sub
End Sub

Sub StrMenu(StrToGrid)
   'Modify By Cheng 2004/03/02
   'strSQL = "SELECT FA01||FA02,DECODE(FA10,'013',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),'020',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)),NA03 FROM FAGENT,NATION WHERE FA01>='" & Left(StrToGrid, 6) & "00' AND FA01<='" & Left(StrToGrid, 6) & "zz' AND fa10=NA01(+) "
   'strSQL = strSQL & " union all SELECT cu01||cu02,DECODE(cu10,'013',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),'020',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),DECODE(cu05,NULL,NVL(cu04,cu06),cu05||' '||cu88||' '||cu89||' '||cu90)),NA03 FROM customer,NATION WHERE cu01>='" & Left(StrToGrid, 6) & "00' AND cu01<='" & Left(StrToGrid, 6) & "zz' AND cu10=NA01(+) "
   strSql = "SELECT FA01||FA02||Decode(FA02,'0','','＊'),SUBSTR(DECODE(FA10,'013',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),'020',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)),1,80),NA03 FROM FAGENT,NATION WHERE FA01>='" & Left(StrToGrid, 6) & "00' AND FA01<='" & Left(StrToGrid, 6) & "zz' AND fa10=NA01(+) "
   strSql = strSql & " union all SELECT cu01||cu02||Decode(CU02,'0','','＊'),SUBSTR(DECODE(cu10,'013',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),'020',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),DECODE(cu05,NULL,NVL(cu04,cu06),cu05||' '||cu88||' '||cu89||' '||cu90)),1,80),NA03 FROM customer,NATION WHERE cu01>='" & Left(StrToGrid, 6) & "00' AND cu01<='" & Left(StrToGrid, 6) & "zz' AND cu10=NA01(+) "
   'End
   'Add By Sindy 98/03/20
   strSql = strSql & " union  SELECT PCU01||PCU02||Decode(PCU02,'0','','＊'),NVL(PCU08,DECODE(PCU03,NULL,PCU07,PCU03||' '||PCU04||' '||PCU05||' '||PCU06)),NA03 FROM PotCustomer,Nation WHERE PCU01>='" & Left(StrToGrid, 6) & "00' AND PCU01<='" & Left(StrToGrid, 6) & "zz'   AND NA01(+)=PCU09"
   strSql = strSql & " union  SELECT POC01||POC02||Decode(POC02,'0','','＊'),POC03,NA03 FROM PotCustomer1,Nation WHERE POC01>='" & Left(StrToGrid, 6) & "00' AND POC01<='" & Left(StrToGrid, 6) & "zz'   AND NA01(+)=POC04"
   '傳入R1時找出相關的X
   strSql = strSql & " union  SELECT cu01||cu02||Decode(CU02,'0','','＊'),SUBSTR(DECODE(cu10,'013',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),'020',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),DECODE(cu05,NULL,NVL(cu04,cu06),cu05||' '||cu88||' '||cu89||' '||cu90)),1,80),NA03 " & _
                                                    "From CUSTOMER, PotCustomer1, Nation " & _
                                               "WHERE CU10=NA01(+) " & _
                                                    "AND POC01>='" & Left(StrToGrid, 6) & "00' AND POC01<='" & Left(StrToGrid, 6) & "zz' " & _
                                                    "AND CU01>=(substr(POC16,1,6)||'00') AND CU01<=(substr(POC16,1,6)||'zz') " & _
                                                    "AND POC16 is not null "
   '找出R1的關係企業
   strSql = strSql & " union  SELECT POC01||POC02||Decode(POC02,'0','','＊'),POC03,NA03 " & _
                                                    "From PotCustomer1, Nation " & _
                                                "WHERE NA01(+)=POC04 " & _
                                                     "AND POC16>='" & Left(StrToGrid, 6) & "00' AND POC16<='" & Left(StrToGrid, 6) & "zz' " & _
                                                     "AND POC16 is not null "
   '傳入R1時找出相關的R
   strSql = strSql & " union  SELECT PCU01||PCU02||Decode(PCU02,'0','','＊'),NVL(PCU08,DECODE(PCU03,NULL,PCU07,PCU03||' '||PCU04||' '||PCU05||' '||PCU06)),NA03 " & _
                                                    "From PotCustomer, Nation, PotCustomer1 " & _
                                               "WHERE NA01(+)=PCU09 " & _
                                                    "AND POC01>='" & Left(StrToGrid, 6) & "00' AND POC01<='" & Left(StrToGrid, 6) & "zz' " & _
                                                    "AND PCU47>=(substr(POC16,1,6)||'00') AND PCU47<=(substr(POC16,1,6)||'zz') " & _
                                                    "AND POC16 is not null AND PCU47 is not null "
   '98/03/19 End
   'Add By Sindy 2009/06/24
   '傳入R時找出相關的X
   strSql = strSql & " union  SELECT cu01||cu02||Decode(CU02,'0','','＊'),SUBSTR(DECODE(cu10,'013',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),'020',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),DECODE(cu05,NULL,NVL(cu04,cu06),cu05||' '||cu88||' '||cu89||' '||cu90)),1,80),NA03 " & _
                                                    "From CUSTOMER, PotCustomer, Nation " & _
                                               "WHERE CU10=NA01(+) " & _
                                                    "AND PCU01>='" & Left(StrToGrid, 6) & "00' AND PCU01<='" & Left(StrToGrid, 6) & "zz' " & _
                                                    "AND CU01>=(substr(PCU47,1,6)||'00') AND CU01<=(substr(PCU47,1,6)||'zz') " & _
                                                    "AND PCU47 is not null "
   '傳入R時找出相關的Y
   strSql = strSql & " union  SELECT FA01||FA02||Decode(FA02,'0','','＊'),NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),NA03 " & _
                                                    "From Fagent, PotCustomer, Nation " & _
                                                "WHERE NA01(+)=FA10 " & _
                                                     "AND PCU01>='" & Left(StrToGrid, 6) & "00' AND PCU01<='" & Left(StrToGrid, 6) & "zz' " & _
                                                     "AND FA01>=(substr(PCU47,1,6)||'00') AND FA01<=(substr(PCU47,1,6)||'zz') " & _
                                                     "AND PCU47 is not null "
   '找出R的關係企業
   strSql = strSql & " union  SELECT PCU01||PCU02||Decode(PCU02,'0','','＊'),NVL(PCU08,DECODE(PCU03,NULL,PCU07,PCU03||' '||PCU04||' '||PCU05||' '||PCU06)),NA03 " & _
                                                    "From PotCustomer, Nation " & _
                                               "WHERE NA01(+)=PCU09 " & _
                                                    "AND PCU47>='" & Left(StrToGrid, 6) & "00' AND PCU47<='" & Left(StrToGrid, 6) & "zz' " & _
                                                    "AND PCU47 is not null "
   '傳入R時找出相關的R1
   strSql = strSql & " union  SELECT POC01||POC02||Decode(POC02,'0','','＊'),POC03,NA03 " & _
                                                    "From PotCustomer1, Nation, PotCustomer " & _
                                               "WHERE NA01(+)=POC04 " & _
                                                    "AND PCU01>='" & Left(StrToGrid, 6) & "00' AND PCU01<='" & Left(StrToGrid, 6) & "zz' " & _
                                                    "AND POC16>=(substr(PCU47,1,6)||'00') AND POC16<=(substr(PCU47,1,6)||'zz') " & _
                                                    "AND PCU47 is not null AND POC16 is not null "
   '2009/06/24 End
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 Then
       adoRecordset.MoveFirst
       'cnnConnection.Execute "delete from r100114 where id='" & strUserNum & "' " Mark by Amy 2015/03/27
       Do While adoRecordset.EOF = False
       
       strSql = "INSERT INTO R100114 values ('"
       If Not IsNull(adoRecordset.Fields(0)) Then
           strSql = strSql + ChgSQL(adoRecordset.Fields(0)) + "','"
       Else
           strSql = strSql + "','"
       End If
       If Not IsNull(adoRecordset.Fields(1)) Then
           strSql = strSql & ChgSQL(adoRecordset.Fields(1)) + "','"
       Else
           strSql = strSql + "','"
       End If
       If Not IsNull(adoRecordset.Fields(2)) Then
           strSql = strSql + ChgSQL(adoRecordset.Fields(2)) + "','" & strUserNum & "')"
       Else
           strSql = strSql + "','" & strUserNum & "')"
       End If
       cnnConnection.Execute strSql
       adoRecordset.MoveNext
       Loop
   Else
   End If
   CheckOC
End Sub

Sub StrMenu1()
    Dim k As Integer 'Add by Amy 2019/10/05
    
   Screen.MousePointer = vbHourglass
   '92.10.15 MODIFY BY SONIA
   'strSQL = "SELECT '' AS V,R07001 AS 編號,R07002 AS 名稱,R07003 AS 國籍 FROM R100114 ORDER BY 編號"
   'edit by nickc 2005/12/06
   ' strSQL = "SELECT '' AS V,R07001 AS 編號,R07002 AS 名稱,R07003 AS 國籍,CU80 AS 狀態,CU79 AS 備註 FROM R100114,CUSTOMER where SUBSTR(R07001,1,1)='X' AND SUBSTR(R07001,1,8)=CU01(+) AND SUBSTR(R07001,9,1)=CU02(+)"
   'strSQL = strSQL & "UNION ALL SELECT '' AS V,R07001 AS 編號,R07002 AS 名稱,R07003 AS 國籍,FA69 AS 狀態,FA29 AS 備註 FROM R100114,FAGENT where SUBSTR(R07001,1,1)='Y' AND SUBSTR(R07001,1,8)=FA01(+) AND SUBSTR(R07001,9,1)=FA02(+) ORDER BY 編號"
   'Modify by Amy 2013/12/10 +智權人/申請國家/總收文號/案件性質/收文日
   'Modify by Amy 2015/05/12 智權人 抓ST02
   'Added by Lydia 2017/12/05 改由啟用日控制
   If strSrvDate(1) < 國外部關聯企業啟用日 Then
        'Modify by Amy 2019/10/05 +4個''->關聯編號/名稱/關係/說明 避免加欄位困難
        'Modified by Lydia 2020/05/07 +'00' as R11401
         strSql = "SELECT '' AS V,R07001||decode(cu111,'Y','$','') AS 編號,R07002 AS 名稱,R07003 AS 國籍,ST02 as 智權人員,CU80 AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日,'' as 關聯編號,'' as 關聯名稱,'' as 關聯關係,'' as 關聯說明,'00' as R11401 FROM R100114,CUSTOMER,Staff where id='" & strUserNum & "' And SUBSTR(R07001,1,1)='X' AND SUBSTR(R07001,1,8)=CU01(+) AND SUBSTR(R07001,9,1)=CU02(+) AND CU13=ST01(+) "
        'Modify by Amy 2019/10/05 原:Union All 把All  拿掉 ex:X29973 有兩筆(一筆為更名)->兩筆勾選->按「關係企業」->不應出現四筆
        strSql = strSql & "UNION  SELECT '' AS V,R07001||decode(fa77,'Y','$','') AS 編號,R07002 AS 名稱,R07003 AS 國籍,'' as 智權人員,FA69 AS 狀態,FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日,'' as 關聯編號,'' as 關聯名稱,'' as 關聯關係,'' as 關聯說明,'00' as R11401 FROM R100114,FAGENT where id='" & strUserNum & "' And SUBSTR(R07001,1,1)='Y' AND SUBSTR(R07001,1,8)=FA01(+) AND SUBSTR(R07001,9,1)=FA02(+) "
        '92.10.15 END
        'Add By Sindy 98/03/20
        'Modify by Amy 2015/05/12 智權人 抓ST02
        'Modify by Amy 2020/03/16 智權人員 原:st02 ,因開發人員可能多人
        strSql = strSql & "UNION  SELECT '' AS V,R07001 AS 編號,R07002 AS 名稱,R07003 AS 國籍,pcu38 as 智權人員,PCU39 AS 狀態,PCU40 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日,'' as 關聯編號,'' as 關聯名稱,'' as 關聯關係,'' as 關聯說明,'00' as R11401 FROM R100114,POTCUSTOMER,Staff where id='" & strUserNum & "' AND SUBSTR(R07001,1,1)='R' AND SUBSTR(R07001,1,8)=PCU01 AND SUBSTR(R07001,9,1)=PCU02 and substr(LTrim(PCU38),1,5)=ST01(+) "
        strSql = strSql & "UNION  SELECT '' AS V,R07001 AS 編號,R07002 AS 名稱,R07003 AS 國籍,poc13 as 智權人員,POC14 AS 狀態,POC15 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日,'' as 關聯編號,'' as 關聯名稱,'' as 關聯關係,'' as 關聯說明,'00' as R11401 FROM R100114,POTCUSTOMER1,Staff where id='" & strUserNum & "' AND SUBSTR(R07001,1,1)='R' AND SUBSTR(R07001,1,8)=POC01 AND SUBSTR(R07001,9,1)=POC02 and POC13=ST01(+) "
        'strSql = strSql & "ORDER BY 編號" 'Remove by Amy 2019/10/05 +活化客戶
        'end 2020/03/16
        'end 2015/05/12
        '98/03/20 End
   Else
        'Added by Lydia 2017/02/14 抓關聯企業改成模組,暫存R100114_1
        'Modified by Lydia 2020/05/07 +R11401
        strSql = "SELECT '' AS V,R11402 AS 編號,R11403 AS 名稱,NVL(NA03,R11405) AS 國籍 ,ST02 AS 智權人員,R11407 AS 狀態,R11408 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日," & _
                 "R11409 AS 關聯編號,DECODE(SUBSTR(R11409,1,1)," & _
                 "'X',DECODE(SIGN(INSTR('000,001,002,003,004,005,006,007,008,009,013,020',C1.CU10)),0,DECODE(C1.CU05,NULL,NVL(C1.CU04,C1.CU06),C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90),NVL(C1.CU04,DECODE(C1.CU05,NULL,C1.CU06,C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)))," & _
                 "'Y',DECODE(SIGN(INSTR('000,001,002,003,004,005,006,007,008,009,013,020',F1.FA10)),0,DECODE(F1.FA05,NULL,NVL(F1.FA04,F1.FA06),F1.FA05||' '||F1.FA63||' '||F1.FA64||' '||F1.FA65),NVL(F1.FA04,DECODE(F1.FA05,NULL,F1.FA06,F1.FA05||' '||F1.FA63||' '||F1.FA64||' '||F1.FA65)))" & _
                 ",R11409) AS 關聯名稱," & _
                 "R11410 AS 關聯關係, R11411 AS 關聯說明,R11401 FROM R100114_1,STAFF,NATION,CUSTOMER C1,FAGENT F1 " & _
                 "WHERE ID='" & strUserNum & "' AND FORMID='" & UCase(Me.Name) & "' AND R11406=ST01(+) AND R11405=NA01(+) " & _
                 "AND SUBSTR(R11409,1,8)=C1.CU01(+) AND '0'=C1.CU02(+) AND SUBSTR(R11409,1,8)=F1.FA01(+) AND '0'=F1.FA02(+) "
        'strSql = strSql & "ORDER BY R11401,R11402,R11409 " 'Remove by Amy 2019/10/05 +活化客戶
        'end 2017/02/14
   End If
   'end 2017/12/05
   
   'Added by Amy 2019/10/05 +活化客戶
   'Modified by Lydia 2020/05/07 重新整理SQL
   'strSql = "Select X.*,Decode(Ocu01,null, '',NVL(Ocu03,0)) as OCU03 from (" & strSql & ") X, OldCustomer Where substr(編號,1,8)= ocu01(+) "
   'Modify by Amy 2023/08/23 更名OCU03=>待活化客戶;增加ORGN
   strSql = "Select X.V, X.編號, X.名稱, X.國籍, X.智權人員, X.狀態, X.備註, X.申請國家, X.總收文號, X.案件性質, X.收文日, X.關聯編號, X.關聯名稱, X.關聯關係, X.關聯說明, " & _
               "'' as ORGN, Decode(Ocu01,null, '',NVL(Ocu03,0)) as 待活化客戶 from (" & strSql & ") X, OldCustomer Where substr(編號,1,8)= ocu01(+) "
   If strSrvDate(1) < 國外部關聯企業啟用日 Then
        strSql = strSql & " ORDER BY 編號"
   Else
        'Modified by Lydia 2020/05/07 重新整理SQL
        'strSql = strSql & " ORDER BY R11401,R11402,R11409 "
        strSql = strSql & " ORDER BY R11401, 編號, 關聯編號"
   End If
   'end 2019/10/05

   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   '911029 nick edit
   If adoRecordset.RecordCount <> 0 Then
       Set grdDataList.Recordset = adoRecordset
       'Modify by Amy 2023/08/28 原程式搬至SetDataListWidth
        SetDataListWidth (True)
   End If
   CheckOC
   
   'Add by Amy 2019/10/05 +所有顏色顯示
   grdDataList.Visible = False
   If grdDataList.Rows > 0 Then
        For i = 1 To grdDataList.Rows - 1
            grdDataList.row = i
            grdDataList.col = 1
            grdDataList.CellForeColor = &H0   '字黑色 ex:查儀大會整個變黑
            'Modify by Amy 2023/08/28 變色改共用函數
            'Modify by Amy 2023/09/26 依狀態更新智權人員改為共用函數
            Call UpdQuerySales(Me.Name, grdDataList, strField)
            'end 2023/09/26
            Call SetMSGridColorQCus(0, Me.Name, grdDataList, strField, IIf(Check3.Value = vbChecked, True, False))
            'end 2023/08/28
        Next i
   End If
   'end 2019/10/05
   'SetDataListWidth 'Remove by Lydia 2017/02/14
   '若只有一筆資料 , 則直接設定為點選此筆資料
   'Modify by Amy 2023/08/28 原程式改成共用SetGridOneData,避免有沒改到的
   cmdOK(6).BackColor = &H8000000F
   Call SetGridOneData
   'end 2023/08/28
   grdDataList.Visible = True
   Screen.MousePointer = vbDefault
End Sub

'Modify by Amy 2022/08/04 名稱查詢語法改至共用Function,並整理程式
Private Sub cmdSearch_Click()
    Dim stSQLa As String, stSqlB As String, stSqlC As String, stSqlD As String, stSqlE As String
    Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
    Dim strSwhSQL1 As String, strSwhSQL2 As String, strSubSQL1 As String, strSubSQL2 As String
    Dim strCheckWay As String, strNo As String, Str01 As String, strFields As String
    Dim IsDevelop As Boolean
    Dim strRtnVal As String 'Add by Amy 2023/08/17
On Error GoTo ErrHnd
    bolPrint = False '先設定無對造
    StrToPrint = ""
   
    '編號
    If Option1(0).Value = True Then
        If Len(Trim(txt1(0))) = 0 Then
            s = MsgBox("編號不可空白", , "USER 輸入資料錯誤")
            txt1(0).SetFocus
            Exit Sub
        End If
    '名稱
    Else
        If Option1(1).Value = True Then
            If Len(Trim(txtFM2(0))) = 0 Then
                s = MsgBox("名稱不可空白", , "USER 輸入資料錯誤")
                txtFM2(0).SetFocus
                Exit Sub
            End If
        Else
            If Option1(2).Value = True Then
                If Len(Trim(txt1(9))) = 0 Then
                    s = MsgBox("國籍不可空白", , "USER 輸入資料錯誤")
                    txt1(9).SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If
   
   'E-mail
    If Option1(3).Value = True Then
        If Len(Trim(txt1(10))) = 0 Then
            s = MsgBox("條件不可空白", , "輸入條件錯誤")
            txt1(10).SetFocus
            Exit Sub
        End If
    End If
    
   'Add by Amy 2023/08/17 屬於查詢置換字彈訊息
   If Option1(1).Value = True Then
      If ChkQuryChangetxt(txtFM2(0), strRtnVal) = True Then
         frm100137_1.Caption = "訊息"
         frm100137_1.txtOrg = txtFM2(0)
         frm100137_1.txtChg = strRtnVal
         frm100137_1.Show vbModal
      End If
   End If
   
    ClearQueryLog (Me.Name) '清除查詢印表記錄檔欄位
    Screen.MousePointer = vbHourglass
    grdDataList.Clear
    grdDataList.Rows = 2
    SetDataListWidth
   
    'Modify by Amy 2022/08/19 +OrgN
    strFields = ",'' AS 關聯編號,'' AS 關聯名稱,'' AS 關聯關係,'' AS 關聯說明,'' AS OrgN "
   
    '若國籍 台灣/香港/大陸 名稱抓中-->英-->日, 否則抓英-->中-->日
    stSQLa = "Decode(Sign(InStr('000,001,002,003,004,005,006,007,008,009,013,020',FA10)),1,Nvl(FA04,Decode(FA05,null,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),Decode(FA05,null,Nvl(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 名稱,"
    stSqlB = "Decode(Sign(InStr('000,001,002,003,004,005,006,007,008,009,013,020',CU10)),1,Nvl(CU04,Decode(CU05,null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),Decode(CU05,null,Nvl(CU04,CU06),CU05||' '||CU88||' '||CU89||' '||CU90 )) as 名稱,"
    stSqlC = "Decode(Sign(InStr('000,001,002,003,004,005,006,007,008,009,013,020',PCU09)),1,Nvl(PCU08,Decode(PCU03,null,PCU07,PCU03||' '||PCU04||' '||PCU05||' '||PCU06)),Decode(PCU03,null,Nvl(PCU07,PCU08),PCU03||' '||PCU04||' '||PCU05||' '||PCU06)) as 名稱,"
    stSqlD = "Decode(Sign(InStr('000,001,002,003,004,005,006,007,008,009,013,020',POC04)),1,Nvl(POC03,Decode(POC23,null,POC28,POC23||' '||POC24||' '||POC25||' '||POC26)),Decode(POC23,null,Nvl(POC03,POC28),POC23||' '||POC24||' '||POC25||' '||POC26)) as 名稱,"
    stSqlE = "Decode(Sign(InStr('000,001,002,003,004,005,006,007,008,009,013,020',NT08)),1,Nvl(NT02,Decode(NT03,null,NT07,NT03||' '||NT04||' '||NT05||' '||NT06)),Decode(NT03,null,Nvl(NT02,NT07),NT03||' '||NT04||' '||NT05||' '||NT06)) as 名稱,"

    If Option1(0).Value = True Then
'*** 編號查詢 ***
        '潛在客戶
        If UCase(Left(Trim(txt1(0)), 1)) = "R" Then
            strSql = "Select ' ' AS V ,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號," & stSqlC & "NA03 AS 國籍,PCU38 AS 智權人員,PCU39 AS 狀態,PCU40 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From POTCUSTOMER,NATION,STAFF Where PCU09=NA01(+) And PCU01='" & Left(GetNewFagent(Trim(txt1(0))), 8) & "' And SubStr(LTrim(PCU38),1,5)=ST01(+) "
            strSql = strSql & " Union All " & _
                         "Select ' ' AS V ,POC01||POC02||Decode(POC02,'0','','＊') AS 編號," & stSqlD & "NA03 AS 國籍,POC13 AS 智權人員,POC14 AS 狀態,POC15 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From POTCUSTOMER1,NATION,STAFF Where POC04=NA01(+) And POC01='" & Left(GetNewFagent(Trim(txt1(0))), 8) & "' And POC13=ST01(+) "
        Else
            strSql = "Select ' ' AS V,FA01||FA02||Decode(FA02,'0','','＊')||Decode(fa77,'Y','$','') AS 編號," & stSQLa & "NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From FAGENT,NATION Where FA01='" & Left(GetNewFagent(Trim(txt1(0))), 8) & "' And fa10=na01(+) "
            strSql = strSql & " Union All " & _
                        "Select ' ' AS V,cu01||cu02||Decode(CU02,'0','','＊')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','●','') AS 編號," & stSqlB & "NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From customer,NATION,STAFF Where cu01='" & Left(GetNewFagent(Trim(txt1(0))), 8) & "' And cu10=na01(+) And CU13=ST01(+) "
            strSql = strSql & " Union All " & _
                        "Select ' ' AS V,NT01||Decode(NT21,null,'⊕','') AS 編號," & stSqlE & "NA03 AS 國籍,ST02 AS 智權人員,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From notagent,nation,STAFF Where nt08=na01(+) And nt01='" & IIf(Len(Trim(txt1(0))) >= 3, Trim(txt1(0)), Right("000" & Trim(txt1(0)), 3)) & "' And nt18=ST01(+) "
            'Add by Amy 2023/12/11 +風險檢查對象
            strSql = strSql & " Union All " & GetSearchRiskChkSql(1, Me.Name, Trim(txt1(0)))
        End If
        pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & Trim(txt1(0))
    ElseIf Option1(1).Value = True Then
'*** 名稱查詢 ***
        '模糊比對
        If Option3(0).Value = False Then
            strCheckWay = ">0"
            pub_QL05 = pub_QL05 & ";" & Option3(1).Caption
        '字首比對
        Else
            strCheckWay = "=1"
            pub_QL05 = pub_QL05 & ";" & Option3(0).Caption
        End If
        '對造
        strSQL1 = " And CP01 IN (" & SQLGrpStr(GetGroupKindByTwo, 2) & ") "
        strSQL2 = " And CP01 IN (" & SQLGrpStr("", 1) & ") "
        StrSQL3 = " And CP01 IN (" & SQLGrpStr("", 3) & ") "
        StrSQL4 = " And CP01 IN (" & SQLGrpStr("", 4) & ") "
        strSQL5 = " And CP01 IN (" & SQLGrpStr("", 5) & ") "
        '含投資法務開拓
        If Check1.Value = 1 Then IsDevelop = True
        '刪除暫存檔
        cnnConnection.Execute "Delete From R100102_1 Where ID='" & strUserNum & "@" & Me.Name & "' "
        '固定查對造
        strSql = GetSearchNameSql(Me.Name, txtFM2(0), strCheckWay, IsDevelop, True, strSQL1, strSQL2, StrSQL3, StrSQL4, strSQL5)
        pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & Trim(txtFM2(0))
    ElseIf Option1(2).Value = True Then
'*** 國籍查詢 ***
        strSql = "Select ' 'AS V,FA01||FA02||Decode(FA02,'0','','＊')||Decode(fa77,'Y','$','') AS 編號," & stSQLa & "NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From FAGENT,NATION Where InStr(FA10, '" & txt1(9) & "') = 1 And fa10=NA01(+) "
        strSql = strSql & " Union All Select ' ' AS V ,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號," & stSqlC & "NA03 AS 國籍,PCU38 AS 智權人員,PCU39 AS 狀態,PCU40 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From POTCUSTOMER,NATION,Staff Where InStr(PCU09, '" & txt1(9) & "') = 1 And PCU09=NA01(+) And SubStr(LTrim(PCU38),1,5)=ST01(+) "
        strSql = strSql & " Union All Select ' ' AS V ,POC01||POC02||Decode(POC02,'0','','＊') AS 編號," & stSqlD & "NA03 AS 國籍,POC13 AS 智權人員,POC14 AS 狀態,POC15 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From POTCUSTOMER1,NATION,Staff Where InStr(POC04, '" & txt1(9) & "') = 1 And POC04=NA01(+) And POC13=ST01(+) "
        strSql = strSql & " Union All Select ' ' AS V,cu01||cu02||Decode(CU02,'0','','＊')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','●','') AS 編號," & stSqlB & "NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From customer,NATION,Staff Where InStr(CU10, '" & txt1(9) & "') = 1 And cu10=na01(+) And CU13=ST01(+) "
        strSql = strSql & " Union All Select ' ' AS V,NT01||Decode(NT21,null,'⊕','') AS 編號," & stSqlE & "NA03 AS 國籍,ST02 AS 智權人員,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From notagent,nation,Staff Where nt08=na01(+) And InStr(nt08, '" & txt1(9) & "') = 1 And nt18=ST01(+) "
        'Add by Amy 2023/12/11 +風險檢查對象
        strSql = strSql & " Union All " & GetSearchRiskChkSql(3, Me.Name, Trim(txt1(9)))
        pub_QL05 = pub_QL05 & ";" & Option1(2).Caption & txt1(9)
    ElseIf Option1(3).Value = True Then
'*** E-Mail 查詢 ***
        'Modified by Lydia 2024/09/18 +財務副本信箱CU200
        strSql = "Select ' ' AS V,CU01||CU02||Decode(CU02,'0','','＊')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','●','') AS 編號,Nvl(CU04,Decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From CUSTOMER,NATION,Staff  Where (InStr(NLS_Upper(CU20),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')>0 Or InStr(NLS_Upper(CU115),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')>0 or InStr(NLS_Upper(CU116),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')>0  or InStr(NLS_Upper(CU117),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')>0 or InStr(NLS_Upper(CU118),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 or InStr(NLS_Upper(CU200),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 )  And CU10=NA01(+) And CU13=ST01(+) "
        strSql = strSql & " Union All " & _
                    "Select ' ' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,Nvl(PCU08,Decode(PCU03,NULL,PCU07,PCU03||' '||PCU04||' '||PCU05||' '||PCU06)) AS 名稱,NA03 AS 國籍,PCU38 AS 智權人員,PCU39 AS 狀態,PCU40 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From potcustomer,nation,Staff  Where (InStr(NLS_Upper(PCU18),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') >0 ) And PCU09=na01(+) And SubStr(LTrim(PCU38),1,5)=ST01(+) "
        strSql = strSql & " Union All " & _
                    "Select ' ' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號," & stSqlD & "NA03 AS 國籍,POC13 AS 智權人員,POC14 AS 狀態,POC15 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From potcustomer1,nation,Staff  Where (InStr(NLS_Upper(POC09),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') >0 ) And POC04=na01(+) And POC13=ST01(+) "
        'Modified by Lydia 2024/09/18 +財務副本信箱FA134
        strSql = strSql & " Union All " & _
                    "Select ' ' AS V,FA01||FA02||Decode(FA02,'0','','＊')||Decode(fa77,'Y','$','') AS 編號,Nvl(FA04,Decode(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)) AS 名稱,NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From fagent,nation Where (InStr(NLS_Upper(fa16),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')> 0 or InStr(NLS_Upper(fa79),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')> 0  or InStr(NLS_Upper(fa105),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')> 0 or InStr(NLS_Upper(fa80),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')> 0 or InStr(NLS_Upper(fa81),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 Or InStr(NLS_Upper(fa82),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 Or InStr(NLS_Upper(FA134),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 ) And fa10=na01(+)   "
        strSql = strSql & " Union All " & _
                    "Select ' ' AS V,PCC01||'0-'||PCC02 AS 編號,Decode(PCC05,'',PCC03,'',PCC04,PCC05) AS 名稱,' ' AS 國籍,' ' AS 智權人員,' ' AS 狀態,PCC13 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From PotCustCont Where (InStr(NLS_Upper(PCC08),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 )  "
        '含投資法務開拓
        If Check1.Value = 1 Then
            strSql = strSql & " Union All Select ' ' AS V,ecd02||'-'||LPAD(ecd01,6,'0') AS 編號,Nvl(ecd03,'')||Nvl(ecd04,'') AS 名稱,NA03 AS 國籍,' ' AS 智權人員,'投法開拓'||Decode(ecd15,null,null,'-'||ecd15) AS 狀態,ecd16 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From expandcusdetail,expandcusattr,nation Where (InStr(NLS_Upper(ecd13),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 ) And ecd10=na01(+) And ecd02=eca01(+) "
        End If
        'Add By Sindy 2023/8/21 + 電子報特殊名單
        strSql = strSql & " Union All " & _
                    "Select ' ' as V,'電子報特殊名單-'||TBNP09 as 編號,TBNP01 as 名稱,'' as 國籍,'' as 智權人員,TBNP10 as 狀態,'' as 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & _
                    " From TMBulletinNp Where (InStr(NLS_Upper(TBNP01),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 ) And TBNP08='M' "
        '2023/8/21 END
        pub_QL05 = pub_QL05 & ";" & Option1(3).Caption & Trim(txt1(10))
    End If
    '含投資法務開拓
    If Check1.Value = 1 Then
        pub_QL05 = pub_QL05 & ";" & Check1.Caption
    End If
    
    CheckOC
    '名稱
    If Option1(1).Value = True Then
        'Modify by Amy 2022/08/19 因名稱前加找到之中 or 英 or 日欄位,導致同編號無法排於一起 原:Order by Upper(名稱),編號
        'ex: 查 SONN & PARTNER 2筆(Y45656000/1)及投法981-000001,2筆Y編號無法排一起
        strSql = "Select X.*,Decode(Ocu01,null, '',Nvl(Ocu03,0)) AS OCU03 From (" & strSql & ") X, OldCustomer Where SubStr(編號,1,8)= ocu01(+) Order by Upper(OrgN) "
    Else
        strSql = "Select X.*,Decode(Ocu01,null, '',Nvl(Ocu03,0)) AS OCU03 From (" & strSql & ") X, OldCustomer Where SubStr(編號,1,8)= ocu01(+) Order by 編號 "
    End If
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/13 記錄此Form的查詢條件
    If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
        InsertQueryLog (adoRecordset.RecordCount)
        cmdOK(0).Enabled = True
        cmdOK(1).Enabled = True
        cmdOK(2).Enabled = True
        cmdOK(5).Enabled = True
        Set grdDataList.Recordset = adoRecordset
    Else
        InsertQueryLog (0)
        Pub_Can_Copy_Pic = True
        ShowNoData
        Pub_Can_Copy_Pic = False
        cmdOK(0).Enabled = False
        cmdOK(1).Enabled = False
        cmdOK(2).Enabled = False
        cmdOK(5).Enabled = False
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Me.grdDataList.Visible = False
    CheckOC
    SetDataListWidth
    
    'Modify by Amy 2023/03/08 欄位改動態
    With Me.grdDataList
        If .Rows > 0 Then
            For i = 1 To .Rows - 1
                .row = i
                .col = 1
                .CellForeColor = &H0 '字黑色
                'Modify by Amy 2023/08/28 變色改為共用函數(變色設定以共用函數為主-與秀玲確認過)
'                'Add by Amy 2023/01/18 +X 或 Y 編號若無案件顯示▼
'                If Check3.Value = vbChecked And (Left(.Text, 1) = 客戶編號 Or Left(.Text, 1) = 代理人編號) Then
'                    If ChkXYCase(Left(.Text, 9)) = False Then
'                        .Text = .Text & "▼"
'                    End If
'                End If
'                'end 2023/01/18
'                '活化客戶顯示整列黃,若有呆帳編號底為紅其他欄為黃
'                If .TextMatrix(grdDataList.row, GetValue("待活化客戶")) = "0" And Right(.TextMatrix(grdDataList.row, GetValue("編號")), 1) <> "＊" Then
'                    For j = 0 To .Cols - 1
'                        '呆帳
'                        If Right(.Text, 1) = "$" And j = 1 Then
'                            .CellBackColor = &HFF& '紅色
'                        '活化客戶
'                        Else
'                            .col = j
'                            .CellBackColor = vbYellow
'                        End If
'                    Next
'                ElseIf Right(.Text, 1) = "$" Then
'                    .CellBackColor = &HFF&
'                '解散/廢止/撤銷/死亡 顯示黑底粉字
'                ElseIf (Left(.Text, 1) = "Y" Or Left(.Text, 1) = "X" Or Left(.Text, 1) = "R") _
'                  And (.TextMatrix(i, GetValue("狀態")) = "解散" Or .TextMatrix(i, GetValue("狀態")) = "廢止" Or .TextMatrix(i, GetValue("狀態")) = "撤銷" Or .TextMatrix(i, GetValue("狀態")) = "死亡") Then
'                    For j = 0 To .Cols - 1
'                        .col = j
'                        .CellBackColor = &H0 '黑色
'                        .CellForeColor = &HFF00FF '粉紅色
'                    Next j
'                ElseIf Right(.Text, 1) = "⊕" Or .TextMatrix(i, GetValue("狀態")) = "對造" Or .TextMatrix(i, GetValue("狀態")) = "其他相關人" Then
                    'Modify by Amy 2023/09/26 依狀態更新智權人員改為共用函數
                    '對造重抓智權人資料
                    If Me.grdDataList.TextMatrix(i, GetValue("狀態")) = "對造" Or .TextMatrix(i, GetValue("狀態")) = "其他相關人" Then
                        bolPrint = True '有對造資料
'                        strNo = Pub_RplStr(.TextMatrix(i, GetValue("編號")))
'                        StrToPrint = strNo & ","
'                        Str01 = SystemNumber(strNo, 1)
'                        Select Case Str01
'                            Case "FCP", "FG"
'                                .TextMatrix(i, GetValue("智權人員")) = GetPrjSalesNM(PUB_GetFCPSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                            Case "FCL", "LIN"
'                                .TextMatrix(i, GetValue("智權人員")) = GetPrjSalesNM(PUB_GetFCLSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                            Case "FCT"
'                                .TextMatrix(i, GetValue("智權人員")) = GetPrjSalesNM(PUB_GetFCTSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                            Case "S"
'                                If .TextMatrix(i, GetValue("申請國家")) = "000" Then
'                                    .TextMatrix(i, GetValue("智權人員")) = GetPrjSalesNM(PUB_GetFCTSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                                Else
'                                    .TextMatrix(i, GetValue("智權人員")) = GetPrjSalesNM(PUB_GetAKindSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                                End If
'                            Case Else
'                                .TextMatrix(i, GetValue("智權人員")) = GetPrjSalesNM(PUB_GetAKindSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                        End Select
'                        .TextMatrix(i, GetValue("案件性質")) = .TextMatrix(i, GetValue("案件性質")) & PUB_GetRelateCasePropertyName(.TextMatrix(i, GetValue("總收文號")), "1")
'                        '更新智權人員至暫存檔
'                        strExc(0) = "Update R100102_1 Set R021003='" & .TextMatrix(i, GetValue("智權人員")) & "' Where R021014='" & Str01 & "' And R021015='" & SystemNumber(strNo, 2) & "' And R021016='" & SystemNumber(strNo, 3) & "' And R021017='" & SystemNumber(strNo, 4) & "' "
'                        cnnConnection.Execute strExc(0)
                    End If
'                    If Right(.Text, 1) = "⊕" Or .TextMatrix(i, GetValue("狀態")) = "對造" Then
'                        For j = 0 To .Cols - 1
'                            .col = j
'                            .CellBackColor = &H8080FF
'                        Next j
'                    End If
'                'CW03=7.媒介平台,顯示橘色
'                ElseIf Left(.TextMatrix(i, GetValue("編號")), 1) = "平" And .TextMatrix(i, GetValue("案件性質")) = "7" Then
'                     .CellBackColor = &H80FF& '橘色
'                End If
'                '國內外潛在客戶 智權人員欄需重抓資料(可能多筆)
'                If Left(.Text, 1) = "R" Then
'                    .TextMatrix(i, GetValue("智權人員")) = GetDevelopP(.TextMatrix(i, GetValue("智權人員")))
'                End If
                Call UpdQuerySales(Me.Name, grdDataList, strField)
                'end 2023/09/26
                Call SetMSGridColorQCus(0, Me.Name, grdDataList, strField, IIf(Check3.Value = vbChecked, True, False))
                'end 2023/08/28
            Next i
        End If
    End With
    'end 2023/03/08
    
    '若查詢結果僅有一筆資料, 則直接勾選
    'Modify by Amy 2023/08/28 原程式寫至共用
    cmdOK(6).BackColor = &H8000000F
    Call SetGridOneData
    'end 2023/08/28

    Me.grdDataList.Visible = True
    If bolPrint Then
        cmdOK(8).Enabled = True
    Else
        cmdOK(8).Enabled = False
    End If
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrHnd:
    If Err.Number = -2147217900 Then
        MsgBox "輸入的文字無法查詢,請電腦中心協助！"
    Else
        MsgBox Err.Description
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSearch_Click_Old()
''Add By Cheng 2002/07/09
'Dim StrSQLa As String
''910801 nick
'Dim StrSqlB As String
'Dim strSQLc As String
'Dim strSQLD As String 'Add By Sindy 2011/10/11
'Dim strCheckWay As String
'Dim strSQLE As String 'Add By Sindy 2012/3/21
''Add by Amy 2013/12/04
'Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
'Dim strSwhSQL1 As String, strSwhSQL2 As String
'Dim strSubSQL1 As String, strSubSQL2 As String
'Dim strNo As String, Str01 As String
'bolPrint = False '先設定無對造
'StrToPrint = ""
''end 2013/12/04
'Dim strFields As String 'Added by Lydia 2017/02/14 設定關聯代號欄位
'
'   'Modify By Cheng 2002/03/14
'   ''Add By Cheng 2002/01/07
'   'txt1_LostFocus 2
'   If Option1(0).Value = True Then
'       If Len(Trim(txt1(0))) = 0 Then
'           s = MsgBox("編號不可空白", , "USER 輸入資料錯誤")
'           txt1(0).SetFocus
'           Exit Sub
'       End If
'   Else
'       If Option1(1).Value = True Then
'           If Len(Trim(txtFM2(0))) = 0 Then
'               s = MsgBox("名稱不可空白", , "USER 輸入資料錯誤")
'               txtFM2(0).SetFocus
'               Exit Sub
'           End If
'       Else
'           If Option1(2).Value = True Then
'               If Len(Trim(txt1(9))) = 0 Then
'                   s = MsgBox("國籍不可空白", , "USER 輸入資料錯誤")
'                   txt1(9).SetFocus
'                   Exit Sub
'               End If
'           End If
'       End If
'   End If
'
'   'add by Toni 2008/12/03
'   If Option1(3).Value = True Then
'       If Len(Trim(txt1(10))) = 0 Then
'           s = MsgBox("條件不可空白", , "輸入條件錯誤")
'           txt1(10).SetFocus
'           Exit Sub
'       End If
'   End If
'
'   'If Len(Trim(txt1(3))) = 0 Then
'   '    S = MsgBox("查詢別不可空白", , "USER 輸入資料錯誤")
'   '    txt1(3).SetFocus
'   '    Exit Sub
'   'End If
'   'If Len(Trim(txt1(4))) = 0 Or Len(Trim(txt1(5))) = 0 Then
'   '    S = MsgBox("日期區間不可空白", , "USER 輸入資料錯誤")
'   '    If Len(Trim(txt1(5))) = 0 Then txt1(5).SetFocus
'   '    If Len(Trim(txt1(4))) = 0 Then txt1(4).SetFocus
'   '    Exit Sub
'
'   ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/4 清除查詢印表記錄檔欄位
'   Screen.MousePointer = vbHourglass
'   'End If
'   GrdDataList.Clear
'   GrdDataList.Rows = 2
'   SetDataListWidth
'   strFields = ",'' AS 關聯編號,'' AS 關聯名稱,'' AS 關聯關係,'' AS 關聯說明 " 'Added by Lydia 2017/02/14
'
'   'Add By Cheng 2002/07/09
'   '若國籍為"013"或"020"則名稱抓中-->英-->日, 否則抓英-->中-->日
'   'Modified by Lydia 2020/08/21
''   StrSQLa = "DECODE(FA10,'013',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),'020',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 名稱,"
''   StrSqlB = "DECODE(cu10,'013',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),'020',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),DECODE(cu05,NULL,NVL(cu04,cu06),cu05||' '||cu88||' '||cu89||' '||cu90)) as 名稱,"
''   'Add by Morgan 2007/12/14
''   strSQLc = "DECODE(instr('013,020',pcu09),0,decode(pcu03,NULL,nvl(pcu08,pcu07),rtrim(pcu03||' '||pcu04||' '||pcu05||' '||pcu06)),NVL(pcu08,DECODE(pcu03,NULL,pcu07,rtrim(pcu03||' '||pcu04||' '||pcu05||' '||pcu06)))) as 名稱,"
''   'Add By Sindy 2011/10/11
''   strSQLD = "DECODE(instr('013,020',poc04),0,decode(poc23,NULL,nvl(poc03,poc27),rtrim(poc23||' '||poc24||' '||poc25||' '||poc26)),NVL(poc03,DECODE(poc23,NULL,poc27,rtrim(poc23||' '||poc24||' '||poc25||' '||poc26)))) as 名稱,"
''   'Add By Sindy 2012/3/21
''   strSQLE = "DECODE(instr('013,020',nt08),0,decode(nt03,NULL,nvl(nt02,nt07),rtrim(nt03||' '||nt04||' '||nt05||' '||nt06)),NVL(nt02,DECODE(nt03,NULL,nt07,rtrim(nt03||' '||nt04||' '||nt05||' '||nt06)))) as 名稱,"
'   StrSQLa = "Decode(Sign(InStr('000,001,002,003,004,005,006,007,008,009,013,020',FA10)),1,NVL(FA04,Decode(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),Decode(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 名稱,"
'   StrSqlB = "Decode(Sign(InStr('000,001,002,003,004,005,006,007,008,009,013,020',CU10)),1,NVL(CU04,Decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),Decode(CU05,NULL,NVL(CU04,CU06),CU05||' '||CU88||' '||CU89||' '||CU90 )) as 名稱,"
'   strSQLc = "Decode(Sign(InStr('000,001,002,003,004,005,006,007,008,009,013,020',PCU09)),1,Nvl(Pcu08,Decode(Pcu03,Null,Pcu07,Pcu03||' '||Pcu04||' '||Pcu05||' '||Pcu06)),Decode(Pcu03,Null,Nvl(Pcu07,Pcu08),Pcu03||' '||Pcu04||' '||Pcu05||' '||Pcu06)) as 名稱,"
'   strSQLD = "Decode(Sign(InStr('000,001,002,003,004,005,006,007,008,009,013,020',POC04)),1,Nvl(Poc03,Decode(Poc23,Null,Poc28,Poc23||' '||Poc24||' '||Poc25||' '||Poc26)),Decode(Poc23,Null,Nvl(Poc03,Poc28),Poc23||' '||Poc24||' '||Poc25||' '||Poc26)) as 名稱,"
'   strSQLE = "Decode(Sign(InStr('000,001,002,003,004,005,006,007,008,009,013,020',NT08)),1,Nvl(nt02,Decode(nt03,Null,nt07,nt03||' '||nt04||' '||nt05||' '||nt06)),Decode(nt03,Null,Nvl(nt02,nt07),nt03||' '||nt04||' '||nt05||' '||nt06)) as 名稱,"
'   'end 2020/08/21
'
'   'Modify by Amy 2013/10/30 讀取Fagent及Customer的狀態欄時，先檢查FA103或CU142，有值顯示 處理情形的內容，無值才抓原狀態欄位
'   'Modify by Amy 2013/09/30 trim掉空白檢查:編號,名稱,E-Mail
'   'Modify by Morgan 2007/12/14 程式邏輯整理
'   '以編號查詢
'   If Option1(0).Value = True Then
'       'Modify by Amy 2013/12/04 +智權人/申請國家/總收文號/案件性質/收文日
'      'Modify by Morgan 2007/12/14 加可查潛在客戶
'      If UCase(Left(Trim(txt1(0)), 1)) = "R" Then
'         'Modified by Lydia 2017/02/14 + strfields
'         'Modify by Amy 2020/03/16 智權人員 原:st02 ,因開發人員可能多人
'         strSql = "SELECT ' ' AS V ,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號," & strSQLc & "NA03 AS 國籍,pcu38 AS 智權人員,PCU39 AS 狀態,PCU40 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM POTCUSTOMER,NATION,STAFF WHERE PCU09=NA01(+) AND PCU01='" & Left(GetNewFagent(Trim(txt1(0))), 8) & "' and substr(LTrim(PCU38),1,5)=ST01(+) "
'         'Add By Sindy 2011/10/11
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all SELECT ' ' AS V ,PoC01||PoC02||Decode(PoC02,'0','','＊') AS 編號," & strSQLD & "NA03 AS 國籍,poc13 AS 智權人員,PoC14 AS 狀態,PoC15 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM POTCUSTOMER1,NATION,STAFF WHERE PoC04=NA01(+) AND PoC01='" & Left(GetNewFagent(Trim(txt1(0))), 8) & "' and poc13=ST01(+) "
'         'end 2020/03/16
'      Else
'         'edit by nickc 2005/12/06
'         'strSQL = "SELECT '' AS V,FA01||FA02||Decode(FA02,'0','','＊') AS 編號," & StrSQLa & "NA03 AS 國籍,FA69 AS 狀態, FA29 AS 備註 FROM FAGENT,NATION WHERE FA01='" & Left(GetNewFagent(txt1(0)), 8) & "' AND fa10=na01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = "SELECT '' AS V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號," & StrSQLa & "NA03 AS 國籍,' ' as 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM FAGENT,NATION WHERE FA01='" & Left(GetNewFagent(Trim(txt1(0))), 8) & "' AND fa10=na01(+) "
'         'edit by nickc 2005/12/06
'         'strSQL = strSQL & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊') AS 編號," & StrSqlB & "NA03 AS 國籍,CU80 AS 狀態,CU79 AS 備註 FROM customer,NATION WHERE cu01='" & Left(GetNewFagent(txt1(0)), 8) & "' AND cu10=na01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','') AS 編號," & StrSqlB & "NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM customer,NATION,STAFF WHERE cu01='" & Left(GetNewFagent(Trim(txt1(0))), 8) & "' AND cu10=na01(+) AND CU13=ST01(+) "
'         'Add By Sindy 2012/3/21
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號," & strSQLE & "NA03 AS 國籍,ST02 AS 智權人員,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM notagent,nation,STAFF where nt08=na01(+) and nt01='" & IIf(Len(Trim(txt1(0))) >= 3, Trim(txt1(0)), Right("000" & Trim(txt1(0)), 3)) & "' AND nt18=ST01(+) "
'      End If
'      pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & Trim(txt1(0)) 'Add By Sindy 2010/11/4
'   '以名稱查詢
'   ElseIf Option1(1).Value = True Then
'      pub_QL05 = pub_QL05 & ";" & Option1(1).Caption 'Add By Sindy 2010/11/4
'      '模糊比對
'      If Option3(0).Value = False Then
'         strCheckWay = ">0"
'         pub_QL05 = pub_QL05 & ";" & Option3(1).Caption 'Add By Sindy 2010/11/4
'      '字首比對
'      Else
'         strCheckWay = "=1"
'         pub_QL05 = pub_QL05 & ";" & Option3(0).Caption 'Add By Sindy 2010/11/4
'      End If
'      'Add by Amy 2013/12/04
'      strTp(3) = ChgSQL(UCase(Trim(txtFM2(0))))
'      '對造
'      strSQL1 = " AND CP01 IN (" & SQLGrpStr(GetGroupKindByTwo, 2) & ") "
'      strSQL2 = " AND CP01 IN (" & SQLGrpStr("", 1) & ") "
'      StrSQL3 = " AND CP01 IN (" & SQLGrpStr("", 3) & ") "
'      StrSQL4 = " AND CP01 IN (" & SQLGrpStr("", 4) & ") "
'      strSQL5 = " AND CP01 IN (" & SQLGrpStr("", 5) & ") "
'      'end 2013/12/04
''Modify by Amy 2013/12/04 拿掉中英日 +查對造 +智權人,申請國家,總收文號,案件性質,收文日 欄位
''      '以中文名稱查詢
''      If Option2(0).Value = True Then
''         'edit by nickc 2005/12/06
''         'strSQL = "SELECT '' AS V,FA01||FA02||Decode(FA02,'0','','＊') AS 編號,FA04 AS 名稱,NA03 AS 國籍,FA69 AS 狀態, FA29 AS 備註 FROM FAGENT,NATION, (Select Distinct FA01 As A1 From Fagent Where instr(FA04,'" & ChgSQL(txtfm2(0)) & "')>0 ) A WHERE FA01=A.A1 AND fa10=na01(+) "
''         strSql = "SELECT '' AS V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,FA04 AS 名稱,NA03 AS 國籍,Decode(FA103,null,FA69,Decode(FA103,'A','同意抵帳中',Decode(FA103,'B','宣告破產','帳款處理中'))) AS 狀態, FA29 AS 備註 FROM FAGENT,NATION, (Select Distinct FA01 As A1 From Fagent Where instr(FA04,'" & ChgSQL(Trim(txtfm2(0))) & "')" & strCheckWay & " ) A WHERE FA01=A.A1 AND fa10=na01(+) "
''         'End
''         'Add by Morgan 2007/12/14
''         strSql = strSql & " union all SELECT '' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,pcu08 AS 名稱,NA03 AS 國籍,pcu39 AS 狀態, pcu40 AS 備註 FROM POTCUSTOMER,NATION, (Select Distinct pcu01 As A1 From POTCUSTOMER Where instr(pcu08,'" & ChgSQL(Trim(txtfm2(0))) & "')" & strCheckWay & " ) A WHERE pcu01=A.A1 AND pcu09=na01(+) "
''         'end 2007/12/14
''         'edit by nickc 2005/12/06
''         'strSQL = strSQL & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊') AS 編號,cu04 AS 名稱,NA03 AS 國籍,CU80 AS 狀態,CU79 AS 備註 FROM customer,NATION, (Select Distinct CU01 As A1 From Customer Where instr(cu04,'" & ChgSQL(txtfm2(0)) & "')>0 ) A WHERE CU01=A.A1 And cu10=na01(+)"
''         strSql = strSql & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','') AS 編號,cu04 AS 名稱,NA03 AS 國籍,Decode(CU142,null,CU80,Decode(CU142,'A','同意抵帳中',Decode(CU142,'B','宣告破產','帳款處理中'))) AS 狀態,CU79 AS 備註 FROM customer,NATION, (Select Distinct CU01 As A1 From Customer Where instr(cu04,'" & ChgSQL(Trim(txtfm2(0))) & "')" & strCheckWay & " ) A WHERE CU01=A.A1 And cu10=na01(+)"
''         'End
''
''         'Add By Sindy 98/03/20
''         strSql = strSql & " union all SELECT '' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,poc03 AS 名稱,NA03 AS 國籍,poc14 AS 狀態, poc15 AS 備註 FROM POTCUSTOMER1,NATION, (Select Distinct poc01 As A1 From POTCUSTOMER1 Where instr(poc03,'" & ChgSQL(Trim(txtfm2(0))) & "')" & strCheckWay & " ) A WHERE poc01=A.A1 AND poc04=na01(+) "
''         '98/03/20 End
''
''         'Add by Morgan 2007/12/21 加可查聯絡人
''         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,Decode(CU142,null,CU80,Decode(CU142,'A','同意抵帳中',Decode(CU142,'B','宣告破產','帳款處理中'))) AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(txtfm2(0))) & "')" & strCheckWay & ") A,CUSTOMER,NATION WHERE CU10=NA01(+) AND CU01(+)=PCC01 AND CU02='0' "
''         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,PCU39 AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(txtfm2(0))) & "')" & strCheckWay & ") A,potcustomer,nation where pcu09=na01(+)  AND PCU01(+)=PCC01 AND PCU02='0' "
''         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,Decode(FA103,null,FA69,Decode(FA103,'A','同意抵帳中',Decode(FA103,'B','宣告破產','帳款處理中'))) AS 狀態, PCC13 AS 備註 from (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(txtfm2(0))) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' "
''         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,POC14 AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(txtfm2(0))) & "')" & strCheckWay & ") A,potcustomer1,nation where poc04=na01(+)  AND POC01(+)=PCC01 AND POC02='0' "
''         'end 2007/12/21
''         'Add By Sindy 2012/3/21
''         strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號,nt02 AS 名稱,NA03 AS 國籍,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註 from notagent,nation, (Select Distinct nt01 As A1 From notagent Where instr(nt02,'" & ChgSQL(Trim(txtfm2(0))) & "')" & strCheckWay & " ) A where nt08=na01(+) and nt01=A.A1"
''         pub_QL05 = pub_QL05 & ";" & Option2(0).Caption & ";" & Trim(txtfm2(0)) 'Add By Sindy 2010/11/4
''
''      '以英文名稱查詢
''      ElseIf Option2(1).Value = True Then
''         'edit by nickc 2005/12/06
''         'strSQL = "SELECT '' AS V,FA01||FA02||Decode(FA02,'0','','＊') AS 編號,FA05||' '||FA63||' '||FA64||' '||FA65 AS 名稱,NA03 AS 國籍,FA69 AS 狀態, FA29 AS 備註 FROM FAGENT,NATION, (Select Distinct FA01 As A1 From Fagent Where instr(upper(FA05||' '||FA63||' '||FA64||' '||FA65),'" & UCase(ChgSQL(txtfm2(0))) & "')>0 ) A WHERE FA01=A.A1 AND fa10=NA01(+) "
''         'strSQL = strSQL & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊') AS 編號,cu05||' '||cu88||' '||cu89||' '||cu90 AS 名稱,NA03 AS 國籍,CU80 AS 狀態,CU79 AS 備註 FROM customer,NATION, (Select Distinct CU01 As A1 From Customer Where instr(upper(cu05||' '||cu88||' '||cu89||' '||cu90),'" & UCase(ChgSQL(txtfm2(0))) & "')>0 ) A WHERE CU01=A.A1 AND cu10=na01(+)"
''         strSql = "SELECT '' AS V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,FA05||' '||FA63||' '||FA64||' '||FA65 AS 名稱,NA03 AS 國籍,Decode(FA103,null,FA69,Decode(FA103,'A','同意抵帳中',Decode(FA103,'B','宣告破產','帳款處理中'))) AS 狀態, FA29 AS 備註 FROM FAGENT,NATION, (Select Distinct FA01 As A1 From Fagent Where instr(upper(FA05||' '||FA63||' '||FA64||' '||FA65),'" & UCase(ChgSQL(Trim(txtfm2(0)))) & "')" & strCheckWay & " ) A WHERE FA01=A.A1 AND fa10=NA01(+) "
''         strSql = strSql & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','') AS 編號,cu05||' '||cu88||' '||cu89||' '||cu90 AS 名稱,NA03 AS 國籍,Decode(CU142,null,CU80,Decode(CU142,'A','同意抵帳中',Decode(CU142,'B','宣告破產','帳款處理中'))) AS 狀態,CU79 AS 備註 FROM customer,NATION, (Select Distinct CU01 As A1 From Customer Where instr(upper(cu05||' '||cu88||' '||cu89||' '||cu90),'" & UCase(ChgSQL(Trim(txtfm2(0)))) & "')" & strCheckWay & " ) A WHERE CU01=A.A1 AND cu10=na01(+)"
''         'End
''         'Add by Morgan 2007/12/14
''         strSql = strSql & " union all SELECT '' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,rtrim(pcu03||' '||pcu04||' '||pcu05||' '||pcu06) AS 名稱,NA03 AS 國籍,pcu39 AS 狀態, pcu40 AS 備註 FROM POTCUSTOMER,NATION, (Select Distinct pcu01 As A1 From POTCUSTOMER Where instr(upper(pcu03||' '||pcu04||' '||pcu05||' '||pcu06),'" & UCase(ChgSQL(Trim(txtfm2(0)))) & "')" & strCheckWay & " ) A WHERE pcu01=A.A1 AND pcu09=na01(+) "
''         'end 2007/12/14
''         'Add By Sindy 2010/02/12
''         strSql = strSql & " union all SELECT '' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,rtrim(poc23||' '||poc24||' '||poc25||' '||poc26) AS 名稱,NA03 AS 國籍,poc14 AS 狀態, poc15 AS 備註 FROM POTCUSTOMER1,NATION, (Select Distinct poc01 As A1 From POTCUSTOMER1 Where instr(upper(poc23||' '||poc24||' '||poc25||' '||poc26),'" & UCase(ChgSQL(Trim(txtfm2(0)))) & "')" & strCheckWay & " ) A WHERE poc01=A.A1 AND poc04=na01(+) "
''         '2010/02/12 End
''         'Add by Morgan 2007/12/21 加可查聯絡人
''         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,Decode(CU142,null,CU80,Decode(CU142,'A','同意抵帳中',Decode(CU142,'B','宣告破產','帳款處理中'))) AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(txtfm2(0)))) & "')" & strCheckWay & ") A,CUSTOMER,NATION WHERE CU10=NA01(+) AND CU01(+)=PCC01 AND CU02='0' "
''         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,PCU39 AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(txtfm2(0)))) & "')" & strCheckWay & ") A,potcustomer,nation where pcu09=na01(+)  AND PCU01(+)=PCC01 AND PCU02='0' "
''         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,Decode(FA103,null,FA69,Decode(FA103,'A','同意抵帳中',Decode(FA103,'B','宣告破產','帳款處理中'))) AS 狀態, PCC13 AS 備註 from (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(txtfm2(0)))) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' "
''         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,POC14 AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(txtfm2(0)))) & "')" & strCheckWay & ") A,potcustomer1,nation where poc04=na01(+)  AND POC01(+)=PCC01 AND POC02='0' "
''         'end 2007/12/21
''         'Add By Sindy 2012/3/21
''         strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號,NT03||' '||NT04||' '||NT05||' '||NT06 AS 名稱,NA03 AS 國籍,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註 from notagent,nation, (Select Distinct NT01 As A1 From notagent Where instr(upper(NT03||' '||NT04||' '||NT05||' '||NT06),'" & UCase(ChgSQL(Trim(txtfm2(0)))) & "')" & strCheckWay & " ) A where nt08=na01(+) and nt01=A.A1"
''         pub_QL05 = pub_QL05 & ";" & Option2(1).Caption & ";" & Trim(txtfm2(0)) 'Add By Sindy 2010/11/4
''
''      '以日文名稱查詢
''      Else
''         'edit by nickc 2005/12/06
''         'strSQL = "SELECT '' AS V,FA01||FA02||Decode(FA02,'0','','＊') AS 編號,FA06 AS 名稱,NA03 AS 國籍,FA69 AS 狀態, FA29 AS 備註 FROM FAGENT,NATION, (Select Distinct FA01 As A1 From Fagent Where instr(FA06,'" & ChgSQL(txtfm2(0)) & "')>0 ) A WHERE FA01=A.A1 AND fa10=na01(+) "
''         strSql = "SELECT '' AS V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,FA06 AS 名稱,NA03 AS 國籍,Decode(FA103,null,FA69,Decode(FA103,'A','同意抵帳中',Decode(FA103,'B','宣告破產','帳款處理中'))) AS 狀態, FA29 AS 備註 FROM FAGENT,NATION, (Select Distinct FA01 As A1 From Fagent Where instr(FA06,'" & ChgSQL(Trim(txtfm2(0))) & "')" & strCheckWay & " ) A WHERE FA01=A.A1 AND fa10=na01(+) "
''         'End
''         'Add by Morgan 2007/12/14
''         strSql = strSql & " union all SELECT '' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,pcu07 AS 名稱,NA03 AS 國籍,pcu39 AS 狀態, pcu40 AS 備註 FROM POTCUSTOMER,NATION, (Select Distinct pcu01 As A1 From POTCUSTOMER Where instr(pcu07,'" & ChgSQL(Trim(txtfm2(0))) & "')" & strCheckWay & " ) A WHERE pcu01=A.A1 AND pcu09=na01(+) "
''         'end 2007/12/14
''
''         'Add By Sindy 2010/02/12
''         strSql = strSql & " union all SELECT '' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,poc27 AS 名稱,NA03 AS 國籍,poc14 AS 狀態, poc15 AS 備註 FROM POTCUSTOMER1,NATION, (Select Distinct poc01 As A1 From POTCUSTOMER1 Where instr(poc27,'" & ChgSQL(Trim(txtfm2(0))) & "')" & strCheckWay & " ) A WHERE poc01=A.A1 AND poc04=na01(+) "
''         '2010/02/12 End
''
''         'edit by nickc 2005/12/06
''         'strSQL = strSQL & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊') AS 編號,cu06 AS 名稱,NA03 AS 國籍,CU80 AS 狀態,CU79 AS 備註 FROM customer,NATION, (Select Distinct CU01 As A1 From Customer Where instr(cu06,'" & ChgSQL(txtfm2(0)) & "')>0 ) A WHERE CU01=A.A1 AND cu10=na01(+)"
''         strSql = strSql & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','') AS 編號,cu06 AS 名稱,NA03 AS 國籍,Decode(CU142,null,CU80,Decode(CU142,'A','同意抵帳中',Decode(CU142,'B','宣告破產','帳款處理中'))) AS 狀態,CU79 AS 備註 FROM customer,NATION, (Select Distinct CU01 As A1 From Customer Where instr(cu06,'" & ChgSQL(Trim(txtfm2(0))) & "')" & strCheckWay & " ) A WHERE CU01=A.A1 AND cu10=na01(+)"
''         'End
''         'Add by Morgan 2007/12/21 加可查聯絡人
''         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,Decode(CU142,null,CU80,Decode(CU142,'A','同意抵帳中',Decode(CU142,'B','宣告破產','帳款處理中'))) AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(txtfm2(0))) & "')" & strCheckWay & ") A,CUSTOMER,NATION WHERE CU10=NA01(+) AND CU01(+)=PCC01 AND CU02='0' "
''         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,PCU39 AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(txtfm2(0))) & "')" & strCheckWay & ") A,potcustomer,nation where pcu09=na01(+)  AND PCU01(+)=PCC01 AND PCU02='0' "
''         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,Decode(FA103,null,FA69,Decode(FA103,'A','同意抵帳中',Decode(FA103,'B','宣告破產','帳款處理中'))) AS 狀態, PCC13 AS 備註 from (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(txtfm2(0))) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' "
''         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,POC14 AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(txtfm2(0))) & "')" & strCheckWay & ") A,potcustomer1,nation where poc04=na01(+)  AND POC01(+)=PCC01 AND POC02='0' "
''         'end 2007/12/21
''         'Add By Sindy 2012/3/21
''         strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號,NT07 AS 名稱,NA03 AS 國籍,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註 from notagent,nation, (Select Distinct NT01 As A1 From Notagent Where instr(NT07,'" & ChgSQL(Trim(txtfm2(0))) & "')" & strCheckWay & " ) A where nt08=na01(+) and nt01=A.A1"
''         pub_QL05 = pub_QL05 & ";" & Option2(2).Caption & ";" & Trim(txtfm2(0)) 'Add By Sindy 2010/11/4
''      End If
'
'      'Modify by Amy 2014/02/25 對造由下搬上來改語法存至暫存檔
'            'Modified by Lydia 2019/12/26
'            'cnnConnection.Execute "Delete From R100102_1 Where ID='" & strUserNum & "' "
'            cnnConnection.Execute "Delete From R100102_1 Where ID='" & strUserNum & "@" & Me.Name & "' "
'
''Modified by Lydia 2019/12/26 改成共用模組Pub_ProcR100102_1
''            '對造(中)
''            strSubSQL1 = " And InStr(Upper(CP40),'" & strTp(3) & "') " & strCheckWay
''            strSubSQL2 = " And InStr(Upper(CP50),'" & strTp(3) & "') " & strCheckWay
''            strSwhSQL1 = " CP40>' ' "
''            strSwhSQL2 = " CP50>' ' "
''
''      'Modify by Amy 2015/03/27 拿掉對造案件編號符號,+客戶端平台帳號資料
''            '商標
''            strSql = "Insert Into R100102_1 (r021001,r021002,r021003,r021004,r021005,r021006,r021007,r021008,r021009,r021010,r021011,r021012,r021013,r021014,r021015,r021016,r021017,r021018,ID) " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, CP40 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap ,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+)" & _
''                          " and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+)" & _
''                          " and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+) " & strSQL1 & strSubSQL1
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, CP50 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+)" & _
''                          " and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+)" & _
''                          " and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+) " & strSQL1 & strSubSQL2
''            '專利
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP40 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
''                          " and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
''                          " and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+) " & strSQL2 & strSubSQL1
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP50 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
''                          " and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
''                          " and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+) " & strSQL2 & strSubSQL2
''            '法務
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP40 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
''                          " and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
''                          " and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+) " & StrSQL3 & strSubSQL1
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP50 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家 ,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
''                          " and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
''                          " and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+) " & StrSQL3 & strSubSQL2
''            '顧問
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP40 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
''                          " and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
''                          " and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+) " & StrSQL4 & strSubSQL1
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP50 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
''                          " and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
''                          " and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+) " & StrSQL4 & strSubSQL2
''            '服務
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP40 as 名稱,' ' as 智權人,'1' as 狀態,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
''                          " and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
''                          " and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+) " & strSQL5 & strSubSQL1
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP50 as 名稱,' ' as 智權人,'2' as 狀態,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
''                          " and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
''                          " and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+) " & strSQL5 & strSubSQL2
''
''            '對造(英)
''            strSubSQL1 = " And InStr(Upper(CP41),'" & strTp(3) & "') " & strCheckWay
''            strSubSQL2 = " And InStr(Upper(CP51),'" & strTp(3) & "') " & strCheckWay
''            strSwhSQL1 = " CP41>' ' "
''            strSwhSQL2 = " CP51>' ' "
''            '商標
''            strSql = strSql & " Union " & _
''                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, CP41 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                         "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+)" & _
''                          " and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+)" & _
''                          " and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+) " & strSQL1 & strSubSQL1
''            strSql = strSql & " Union " & _
''                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, CP51 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                         "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+)" & _
''                          " and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+)" & _
''                          " and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+) " & strSQL1 & strSubSQL2
''            '專利
''            strSql = strSql & " Union " & _
''                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP41 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                         "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
''                          " and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
''                          " and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+) " & strSQL2 & strSubSQL1
''            strSql = strSql & " Union " & _
''                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP51 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                         "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
''                          " and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
''                          " and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+) " & strSQL2 & strSubSQL2
''            '法務
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP41 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) as 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
''                          " and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
''                          " and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+) " & StrSQL3 & strSubSQL1
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP51 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) as 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
''                          " and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
''                          " and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+) " & StrSQL3 & strSubSQL2
''            '顧問
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP41 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) as 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
''                          " and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
''                          " and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+) " & StrSQL4 & strSubSQL1
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP51 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) as 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
''                          " and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
''                          " and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+) " & StrSQL4 & strSubSQL2
''            '服務
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP41 as 名稱,' ' as 智權人,'1' as 狀態,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) as 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
''                          " and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
''                          " and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+) " & strSQL5 & strSubSQL1
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP51 as 名稱,' ' as 智權人,'2' as 狀態,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) as 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
''                          " and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
''                          " and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+) " & strSQL5 & strSubSQL2
''
''            '對造(日)
''            strSubSQL1 = " And InStr(Upper(CP42),'" & strTp(3) & "') " & strCheckWay
''            strSubSQL2 = " And InStr(Upper(CP52),'" & strTp(3) & "') " & strCheckWay
''            strSwhSQL1 = " CP42>' ' "
''            strSwhSQL2 = " CP52>' ' "
''            '商標
''            strSql = strSql & " Union " & _
''                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, CP42 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                         "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+)" & _
''                          " and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+)" & _
''                          " and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+) " & strSQL1 & strSubSQL1
''            strSql = strSql & " Union " & _
''                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, CP52 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                         "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+)" & _
''                          " and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+)" & _
''                          " and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+) " & strSQL1 & strSubSQL2
''            '專利
''            strSql = strSql & " Union " & _
''                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP42 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                         "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
''                          " and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
''                          " and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+) " & strSQL2 & strSubSQL1
''            strSql = strSql & " Union " & _
''                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP52 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                         "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
''                          " and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
''                          " and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+) " & strSQL2 & strSubSQL2
''            '法務
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP42 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) as 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
''                          " and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
''                          " and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+) " & StrSQL3 & strSubSQL1
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP52 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) as 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
''                          " and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
''                          " and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+) " & StrSQL3 & strSubSQL2
''            '顧問
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP42 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) as 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
''                          " and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
''                          " and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+) " & StrSQL4 & strSubSQL1
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP52 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) as 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
''                          " and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
''                          " and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+) " & StrSQL4 & strSubSQL2
''            '服務
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP42 as 名稱,' ' as 智權人,'1' as 狀態,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) as 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
''                          " and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
''                          " and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+) " & strSQL5 & strSubSQL1
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP52 as 名稱,' ' as 智權人,'2' as 狀態,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) as 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
''                          " and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
''                          " and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+) " & strSQL5 & strSubSQL2
''            cnnConnection.Execute strSql
''
''           '刪除對造與申請人相同資料
''           strSql = "Delete From R100102_1 Where ID='" & strUserNum & "' And (ltrim(rtrim(R021002))=ltrim(rtrim(R021008)) Or ltrim(rtrim(R021002))=ltrim(rtrim(R021009)) " & _
''                       "Or ltrim(rtrim(R021002))=ltrim(rtrim(R021010)) Or ltrim(rtrim(R021002))=ltrim(rtrim(R021011)) Or ltrim(rtrim(R021002))=ltrim(rtrim(R021012))) "
''           cnnConnection.Execute strSql
''      'end 2014/02/25
''
''      'Add by Amy 2014/03/17 將所有商標案InStr(R021014,'T')且案件性質為1202(核准通知)者狀態改為 其他相關人
''      'Modify by Amy 2015/12/03 增加商標案(CFC/S) 案件性質202(申請意見書)及303(延期)者 狀態改為 其他相關人
''      strSql = "Update R100102_1 Set R021004='2' Where (InStr(R021014,'T')>0 or R021014='CFC' or R021014='S') And (R021018='1202' or R021018='202' or R021018='303')"
''      cnnConnection.Execute strSql
''      'end 2014/03/17
''      'Add by Amy 2015/12/03 所有專利案件性質404(延期) 者狀態改為 其他相關人
''      strSql = "Update R100102_1 Set R021004='2' Where (InStr(R021014,'P')>0 or R021014='FG') And R021018='404' "
''      cnnConnection.Execute strSql
''      'end 2015/12/03
'       Call Pub_ProcR100102_1(strUserNum & "@" & Me.Name, strSQL1, strSQL2, StrSQL3, StrSQL4, strSQL5, strTp(3), strCheckWay)
''end 2019/12/26
'
'      '查Fagent 代理人 檔
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = "SELECT '' AS V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,FA04 AS 名稱,NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM FAGENT,NATION, (Select Distinct FA01 As A1 From Fagent Where instr(FA04,'" & ChgSQL(Trim(txtFM2(0))) & "')" & strCheckWay & " ) A WHERE FA01=A.A1 AND fa10=na01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all Select '' AS V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,FA05||' '||FA63||' '||FA64||' '||FA65 AS 名稱,NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM FAGENT,NATION, (Select Distinct FA01 As A1 From Fagent Where instr(upper(FA05||' '||FA63||' '||FA64||' '||FA65),'" & UCase(ChgSQL(Trim(txtFM2(0)))) & "')" & strCheckWay & " ) A WHERE FA01=A.A1 AND fa10=NA01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all Select '' AS V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,FA06 AS 名稱,NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM FAGENT,NATION, (Select Distinct FA01 As A1 From Fagent Where instr(FA06,'" & ChgSQL(Trim(txtFM2(0))) & "')" & strCheckWay & " ) A WHERE FA01=A.A1 AND fa10=na01(+) "
'
'      '查customer 客戶 檔
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','') AS 編號,cu04 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM customer,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where instr(cu04,'" & ChgSQL(Trim(txtFM2(0))) & "')" & strCheckWay & " ) A WHERE CU01=A.A1 And cu10=na01(+) AND CU13=ST01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','') AS 編號,cu05||' '||cu88||' '||cu89||' '||cu90 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM customer,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where instr(upper(cu05||' '||cu88||' '||cu89||' '||cu90),'" & UCase(ChgSQL(Trim(txtFM2(0)))) & "')" & strCheckWay & " ) A WHERE CU01=A.A1 AND cu10=na01(+) AND CU13=ST01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','') AS 編號,cu06 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM customer,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where instr(cu06,'" & ChgSQL(Trim(txtFM2(0))) & "')" & strCheckWay & " ) A WHERE CU01=A.A1 AND cu10=na01(+) AND CU13=ST01(+) "
'
'      'Modify by Amy 2015/04/15 客戶端平台帳號資料
'      'Modified by Lydia 2017/02/14 + strfields
'      'Modify By Sindy 2021/3/25 '' as 案件性質, => CW03 as 案件性質,
'      strSql = strSql & " union all Select ' ' as V,'平台'||CW01 AS 編號, CW12 AS 名稱,'平台' AS 國籍,' ' AS 智權人員,Nvl(CW19,'') AS 狀態,'' AS 備註,' ' as 申請國家,'' as 總收文號,CW03 as 案件性質,CW01 as 收文日" & strFields & " From CustWeb Where InStr(Upper(CW12),'" & strTp(3) & "') " & strCheckWay
'
'      '查potcustomer 國外潛在客戶 檔
'         'Modified by Lydia 2017/02/14 + strfields
'         'Modify by Amy 2020/03/16 智權人員 原:st02 ,因開發人員可能多人
'         strSql = strSql & " union all SELECT '' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,pcu08 AS 名稱,NA03 AS 國籍,pcu38 AS 智權人員,pcu39 AS 狀態, pcu40 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM POTCUSTOMER,NATION,STAFF, (Select Distinct pcu01 As A1 From POTCUSTOMER Where instr(pcu08,'" & ChgSQL(Trim(txtFM2(0))) & "')" & strCheckWay & " ) A WHERE pcu01=A.A1 AND pcu09=na01(+) and substr(LTrim(PCU38),1,5)=ST01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all SELECT '' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,rtrim(pcu03||' '||pcu04||' '||pcu05||' '||pcu06) AS 名稱,NA03 AS 國籍,pcu38 AS 智權人員,pcu39 AS 狀態, pcu40 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM POTCUSTOMER,NATION,STAFF, (Select Distinct pcu01 As A1 From POTCUSTOMER Where instr(upper(pcu03||' '||pcu04||' '||pcu05||' '||pcu06),'" & UCase(ChgSQL(Trim(txtFM2(0)))) & "')" & strCheckWay & " ) A WHERE pcu01=A.A1 AND pcu09=na01(+) and substr(LTrim(PCU38),1,5)=ST01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all SELECT '' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,pcu07 AS 名稱,NA03 AS 國籍,pcu38 AS 智權人員,pcu39 AS 狀態, pcu40 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM POTCUSTOMER,NATION,STAFF, (Select Distinct pcu01 As A1 From POTCUSTOMER Where instr(pcu07,'" & ChgSQL(Trim(txtFM2(0))) & "')" & strCheckWay & " ) A WHERE pcu01=A.A1 AND pcu09=na01(+) and substr(LTrim(PCU38),1,5)=ST01(+) "
'
'      '查potcustomer1 國內潛在客戶 檔
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all SELECT '' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,poc03 AS 名稱,NA03 AS 國籍,poc13 AS 智權人員,poc14 AS 狀態, poc15 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM POTCUSTOMER1,NATION,STAFF, (Select Distinct poc01 As A1 From POTCUSTOMER1 Where instr(poc03,'" & ChgSQL(Trim(txtFM2(0))) & "')" & strCheckWay & " ) A WHERE poc01=A.A1 AND poc04=na01(+) and poc13=ST01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all SELECT '' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,rtrim(poc23||' '||poc24||' '||poc25||' '||poc26) AS 名稱,NA03 AS 國籍,poc13 AS 智權人員,poc14 AS 狀態, poc15 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM POTCUSTOMER1,NATION,STAFF, (Select Distinct poc01 As A1 From POTCUSTOMER1 Where instr(upper(poc23||' '||poc24||' '||poc25||' '||poc26),'" & UCase(ChgSQL(Trim(txtFM2(0)))) & "')" & strCheckWay & " ) A WHERE poc01=A.A1 AND poc04=na01(+) and poc13=ST01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all SELECT '' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,poc27 AS 名稱,NA03 AS 國籍,poc13 AS 智權人員,poc14 AS 狀態, poc15 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM POTCUSTOMER1,NATION,STAFF, (Select Distinct poc01 As A1 From POTCUSTOMER1 Where instr(poc27,'" & ChgSQL(Trim(txtFM2(0))) & "')" & strCheckWay & " ) A WHERE poc01=A.A1 AND poc04=na01(+) and poc13=ST01(+) "
'        'end 2020/03/16
'
'      '查NotAgent 不得代理案件之客戶或代理人 檔
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號,nt02 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM notagent,nation,staff, (Select Distinct nt01 As A1 From notagent Where instr(nt02,'" & ChgSQL(Trim(txtFM2(0))) & "')" & strCheckWay & " ) A where nt08=na01(+) and nt01=A.A1 AND nt18=ST01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號,NT03||' '||NT04||' '||NT05||' '||NT06 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM notagent,nation,staff, (Select Distinct NT01 As A1 From notagent Where instr(upper(NT03||' '||NT04||' '||NT05||' '||NT06),'" & UCase(ChgSQL(Trim(txtFM2(0)))) & "')" & strCheckWay & " ) A where nt08=na01(+) and nt01=A.A1 AND nt18=ST01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號,NT07 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM notagent,nation,staff, (Select Distinct NT01 As A1 From Notagent Where instr(NT07,'" & ChgSQL(Trim(txtFM2(0))) & "')" & strCheckWay & " ) A where nt08=na01(+) and nt01=A.A1 AND nt18=ST01(+) "
'
'      '查聯絡人(中文)
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(txtFM2(0))) & "')" & strCheckWay & ") A,CUSTOMER,NATION,STAFF WHERE CU10=NA01(+) AND CU01(+)=PCC01 AND CU02='0' AND CU13=ST01(+) "
'         'Modify by Amy 2020/03/16 智權人員 原:st02 ,因開發人員可能多人
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,pcu38 AS 智權人員,PCU39 AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(txtFM2(0))) & "')" & strCheckWay & ") A,potcustomer,nation,STAFF where pcu09=na01(+)  AND PCU01(+)=PCC01 AND PCU02='0' AND substr(LTrim(PCU38),1,5)=ST01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,poc13 AS 智權人員,POC14 AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(txtFM2(0))) & "')" & strCheckWay & ") A,potcustomer1,nation,STAFF where poc04=na01(+)  AND POC01(+)=PCC01 AND POC02='0' AND poc13=ST01(+) "
'         'end 2020/03/16
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,''AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(txtFM2(0))) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' "
'
'      '查聯絡人(英文)
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(txtFM2(0)))) & "')" & strCheckWay & ") A,CUSTOMER,NATION,STAFF WHERE CU10=NA01(+) AND CU01(+)=PCC01 AND CU02='0' AND CU13=ST01(+) "
'         'Modify by Amy 2020/03/16 智權人員 原:st02 ,因開發人員可能多人
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,pcu38 AS 智權人員,PCU39 AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(txtFM2(0)))) & "')" & strCheckWay & ") A,potcustomer,nation,STAFF where pcu09=na01(+)  AND PCU01(+)=PCC01 AND PCU02='0' AND substr(LTrim(PCU38),1,5)=ST01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,poc13 AS 智權人員,POC14 AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(txtFM2(0)))) & "')" & strCheckWay & ") A,potcustomer1,nation,STAFF where poc04=na01(+)  AND POC01(+)=PCC01 AND POC02='0' AND poc13=ST01(+) "
'         'end 2020/03/16
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,' ' as 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(txtFM2(0)))) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' "
'
'      '查聯絡人(日文)
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(txtFM2(0))) & "')" & strCheckWay & ") A,CUSTOMER,NATION,STAFF WHERE CU10=NA01(+) AND CU01(+)=PCC01 AND CU02='0' AND CU13=ST01(+) "
'         'Modify by Amy 2020/03/16 智權人員 原:st02 ,因開發人員可能多人
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,pcu38 AS 智權人員,PCU39 AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(txtFM2(0))) & "')" & strCheckWay & ") A,potcustomer,nation,STAFF where pcu09=na01(+)  AND PCU01(+)=PCC01 AND PCU02='0' AND substr(LTrim(PCU38),1,5)=ST01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,poc13 AS 智權人員,POC14 AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(txtFM2(0))) & "')" & strCheckWay & ") A,potcustomer1,nation,STAFF where poc04=na01(+)  AND POC01(+)=PCC01 AND POC02='0' AND poc13=ST01(+) "
'         'end 2020/03/16
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,' ' as 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(txtFM2(0))) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' "
'
'       '抓暫存檔對造 Add by Amy 2014/02/25
'       'Modified by Lydia 2017/02/14 + strfields
'       'Modified by Lydia 2019/12/26 +@+Me.name
'       'Modify by Amy 2020/09/04 +all 因查 金杜 應出現2筆,中/日文都有輸
'        strSql = strSql & " union all select ' ' as V,R021001 AS 編號,R021002 AS 名稱,'' AS 國籍,'' AS 智權人員,Decode(R021004,'1','對造','其他相關人') AS 狀態,'' AS 備註,'' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM R100102_1 Where ID='" & strUserNum & "@" & Me.Name & "' And R021004<3 "
'      'end 2015/03/27
'
'      'Mark by Amy 2014/02/25 往上搬
''      '對造(中)
''            strSubSQL1 = " And InStr(Upper(CP40),'" & strTp(3) & "') " & strCheckWay
''            strSubSQL2 = " And InStr(Upper(CP50),'" & strTp(3) & "') " & strCheckWay
''            strSwhSQL1 = " CP40>' ' "
''            strSwhSQL2 = " CP50>' ' "
''            '商標
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as 編號, CP40 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL1 & _
''                        " Union  Select ' ' as V,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as 編號, CP50 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL2
''            '專利
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as 編號,CP40 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL1 & _
''                        " Union  Select ' ' as V,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as 編號,CP50 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL2
''            '法務
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 編號,CP40 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase,CasePropertyMap Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 編號,CP50 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家 ,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase,CasePropertyMap Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL2
''            '顧問
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') as 編號,CP40 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase,CasePropertyMap Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') as 編號,CP50 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase,CasePropertyMap Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL2
''            '服務
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') as 編號,CP40 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice,CasePropertyMap Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') as 編號,CP50 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice,CasePropertyMap Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL2
''
''            '對造(英)
''            strSubSQL1 = " And InStr(Upper(CP41),'" & strTp(3) & "') " & strCheckWay
''            strSubSQL2 = " And InStr(Upper(CP51),'" & strTp(3) & "') " & strCheckWay
''            strSwhSQL1 = " CP41>' ' "
''            strSwhSQL2 = " CP51>' ' "
''            '商標
''            strSql = strSql & " Union " & _
''                         "Select ' ' as V,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as 編號, CP41 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL1 & _
''                         " Union Select ' ' as V,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as 編號, CP51 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL2
''            '專利
''            strSql = strSql & " Union " & _
''                         "Select ' ' as V,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as 編號,CP41 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL1 & _
''                         " Union Select ' ' as V,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as 編號,CP51 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL2
''            '法務
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 編號,CP41 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase,CasePropertyMap Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 編號,CP51 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase,CasePropertyMap Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL2
''            '顧問
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') as 編號,CP41 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase,CasePropertyMap Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') as 編號,CP51 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase,CasePropertyMap Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL2
''            '服務
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') as 編號,CP41 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice,CasePropertyMap Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') as 編號,CP51 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice,CasePropertyMap Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL2
''
''            '對造(日)
''            strSubSQL1 = " And InStr(Upper(CP42),'" & strTp(3) & "') " & strCheckWay
''            strSubSQL2 = " And InStr(Upper(CP52),'" & strTp(3) & "') " & strCheckWay
''            strSwhSQL1 = " CP42>' ' "
''            strSwhSQL2 = " CP52>' ' "
''            '商標
''            strSql = strSql & " Union " & _
''                         "Select ' ' as V,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as 編號, CP42 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL1 & _
''                         " Union Select ' ' as V,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as 編號, CP52 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL2
''            '專利
''            strSql = strSql & " Union " & _
''                         "Select ' ' as V,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as 編號,CP42 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL1 & _
''                         " Union Select ' ' as V,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as 編號,CP52 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL2
''            '法務
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 編號,CP42 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase,CasePropertyMap Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 編號,CP52 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase,CasePropertyMap Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL2
''            '顧問
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') as 編號,CP42 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase,CasePropertyMap Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') as 編號,CP52 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase,CasePropertyMap Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL2
''            '服務
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') as 編號,CP42 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice,CasePropertyMap Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') as 編號,CP52 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice,CasePropertyMap Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL2
'      'end 2014/02/25
''end 2013/12/04
'
'      ' Add By Sindy 98/02/13 開拓客戶
'      If Check1.Value = 1 Then
'         'Modify by Amy 2013/12/04 +智權人員/申請國家/總收文號/案件性質/收文日
'         'Modify by Amy 2013/09/30 原只檢查ecd11,ecd12卻顯示ecd03,ecd04
'         'strSql = strSql & " union all SELECT ' ' AS V,ecd02||'-'||LPAD(ecd01,6,'0') AS 編號,NVL(ecd03,'')||NVL(ecd04,'') AS 名稱,NA03 AS 國籍,ecd15 AS 狀態,ecd16 AS 備註 from expandcusdetail,expandcusattr,nation,(Select Distinct nvl(ecd01,'')||nvl(ecd02,'') as A1 From expandcusdetail Where instr(ecd11,'" & ChgSQL(Trim(txtfm2(0))) & "') " & strCheckWay & " or instr(ecd12,'" & ChgSQL(Trim(txtfm2(0))) & "') " & strCheckWay & ") A Where ecd10=na01(+) and ecd02=eca01(+) and nvl(ecd01,'')||nvl(ecd02,'')=A.A1 "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all SELECT ' ' AS V,ecd02||'-'||LPAD(ecd01,6,'0') AS 編號,NVL(ecd03,'')||NVL(ecd04,'') AS 名稱,NA03 AS 國籍,' ' AS 智權人員,'投法開拓'||decode(ecd15,null,null,'-'||ecd15) AS 狀態,ecd16 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM expandcusdetail,expandcusattr,nation,(Select Distinct nvl(ecd01,'')||nvl(ecd02,'') as A1 From expandcusdetail Where instr(Upper(ecd03),'" & ChgSQL(UCase(Trim(txtFM2(0)))) & "') " & strCheckWay & " or instr(Upper(ecd04),'" & ChgSQL(UCase(Trim(txtFM2(0)))) & "') " & strCheckWay & ") A Where ecd10=na01(+) and ecd02=eca01(+) and nvl(ecd01,'')||nvl(ecd02,'')=A.A1 "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all SELECT ' ' AS V,ecd02||'-'||LPAD(ecd01,6,'0') AS 編號,NVL(ecd11,'')||NVL(ecd12,'') AS 名稱,NA03 AS 國籍,' ' AS 智權人員,'投法開拓'||decode(ecd15,null,null,'-'||ecd15) AS 狀態,ecd16 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM expandcusdetail,expandcusattr,nation,(Select Distinct nvl(ecd01,'')||nvl(ecd02,'') as A1 From expandcusdetail Where instr(Upper(ecd11),'" & ChgSQL(UCase(Trim(txtFM2(0)))) & "') " & strCheckWay & " or instr(Upper(ecd12),'" & ChgSQL(UCase(Trim(txtFM2(0)))) & "') " & strCheckWay & ") A Where ecd10=na01(+) and ecd02=eca01(+) and nvl(ecd01,'')||nvl(ecd02,'')=A.A1 "
'      End If
'      ' 98/02/13 End
'
'
'   '以國籍查詢
'   ElseIf Option1(2).Value = True Then
'      'Modify by Amy 2013/12/04 +智權人員/申請國家/總收文號/案件性質/收文日
'      'edit by nickc 2005/12/06
'      'strSQL = "SELECT ''AS V,FA01||FA02||Decode(FA02,'0','','＊') AS 編號," & StrSQLa & "NA03 AS 國籍,FA69 AS 狀態, FA29 AS 備註 FROM FAGENT,NATION WHERE INSTR(FA10, '" & txt1(9) & "') = 1 AND fa10=NA01(+) "
'      'Modified by Lydia 2017/02/14 + strfields
'      strSql = "SELECT ''AS V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號," & StrSQLa & "NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM FAGENT,NATION WHERE INSTR(FA10, '" & txt1(9) & "') = 1 AND fa10=NA01(+) "
'      'Add by Morgan 2007/12/14
'      'Modify by Amy 2020/03/16 智權人員 原:st02 ,因開發人員可能多人
'      'Modified by Lydia 2017/02/14 + strfields
'      strSql = strSql & " union all SELECT ' ' AS V ,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號," & strSQLc & "NA03 AS 國籍,pcu38 AS 智權人員,PCU39 AS 狀態,PCU40 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM POTCUSTOMER,NATION,Staff WHERE INSTR(pcu09, '" & txt1(9) & "') = 1 and PCU09=NA01(+) and substr(LTrim(PCU38),1,5)=ST01(+) "
'      'end 2007/12/14
'      'Add By Sindy 2011/10/11
'      'Modified by Lydia 2017/02/14 + strfields
'      strSql = strSql & " union all SELECT ' ' AS V ,PoC01||PoC02||Decode(PoC02,'0','','＊') AS 編號," & strSQLD & "NA03 AS 國籍,poc13 AS 智權人員,PoC14 AS 狀態,PoC15 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM POTCUSTOMER1,NATION,Staff WHERE INSTR(poc04, '" & txt1(9) & "') = 1 and PoC04=NA01(+) and poc13=ST01(+) "
'      'end 2020/03/16
'
'      'edit by nickc 2005/12/06
'      'strSQL = strSQL & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊') AS 編號," & StrSqlB & "NA03 AS 國籍,CU80 AS 狀態,CU79 AS 備註 FROM customer,NATION WHERE INSTR(CU10, '" & txt1(9) & "') = 1 AND cu10=na01(+)"
'      'Modified by Lydia 2017/02/14 + strfields
'      strSql = strSql & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','') AS 編號," & StrSqlB & "NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM customer,NATION,Staff WHERE INSTR(CU10, '" & txt1(9) & "') = 1 AND cu10=na01(+) AND CU13=ST01(+) "
'      'Add By Sindy 2012/3/21
'      'Modified by Lydia 2017/02/14 + strfields
'      strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號," & strSQLE & "NA03 AS 國籍,ST02 AS 智權人員,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM notagent,nation,Staff where nt08=na01(+) and INSTR(nt08, '" & txt1(9) & "') = 1 AND nt18=ST01(+) "
'      pub_QL05 = pub_QL05 & ";" & Option1(2).Caption & txt1(9) 'Add By Sindy 2010/11/4
'
'   'E-Mail  add by Toni 2008/12/03
'   ElseIf Option1(3).Value = True Then
'        'Modified by Lydia 2017/02/14 + strfields
'        strSql = "SELECT ' ' as V,CU01||CU02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','')||decode(cu121,'Y','●','') AS 編號,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM CUSTOMER,NATION,Staff  Where (instr(NLS_Upper(CU20),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')>0 Or instr(NLS_Upper(CU115),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')>0 or instr(NLS_Upper(CU116),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')>0  or instr(NLS_Upper(CU117),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')>0 or instr(NLS_Upper(CU118),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 )  and CU10=NA01(+) AND CU13=ST01(+) "
'
'        'Modify by Amy 2020/03/16 智權人員 原:st02 ,因開發人員可能多人
'        'Modified by Lydia 2017/02/14 + strfields
'        strSql = strSql & " union all SELECT ' ' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,NVL(PCU08,DECODE(PCU03,NULL,PCU07,PCU03||' '||PCU04||' '||PCU05||' '||PCU06)) AS 名稱,NA03 AS 國籍,pcu38 AS 智權人員,PCU39 AS 狀態,PCU40 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM potcustomer,nation,Staff  Where (instr(NLS_Upper(pcu18),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') >0 ) and pcu09=na01(+) and substr(LTrim(PCU38),1,5)=ST01(+) "
'        'Add By Sindy 98/03/20
'        'Modified by Lydia 2017/02/14 + strfields
'        strSql = strSql & " union all SELECT ' ' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號," & strSQLD & "NA03 AS 國籍,poc13 AS 智權人員,POC14 AS 狀態,POC15 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM potcustomer1,nation,Staff  Where (instr(NLS_Upper(poc09),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') >0 ) and poc04=na01(+) and poc13=ST01(+) "
'        '98/03/20 End
'        'end 2020/03/16
'        'Modified by Lydia 2017/02/14 + strfields
'        'Modified by Lydia 2018/07/20 +FA105 財務信箱(CF)
'        'strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)) AS 名稱,NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM fagent,nation Where (instr(NLS_Upper(fa16),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')> 0 or instr(NLS_Upper(fa79),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')> 0 or instr(NLS_Upper(fa80),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')> 0 or instr(NLS_Upper(fa81),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 Or InStr(NLS_Upper(fa82),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 )   and  fa10=na01(+)   "
'        strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)) AS 名稱,NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM fagent,nation " & _
'                     "Where (instr(NLS_Upper(fa16),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')> 0 or instr(NLS_Upper(fa79),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')> 0  or instr(NLS_Upper(fa105),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')> 0 or instr(NLS_Upper(fa80),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')> 0 or instr(NLS_Upper(fa81),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 Or InStr(NLS_Upper(fa82),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 )   and  fa10=na01(+)   "
'        'Modified by Lydia 2017/02/14 + strfields
'        strSql = strSql & " union all SELECT ' ' AS V,PCC01||'0-'||PCC02 AS 編號,Decode(PCC05,'',PCC03,'',PCC04,PCC05) AS 名稱,' ' AS 國籍,' ' as 智權人員,' ' AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM PotCustCont Where (instr(NLS_Upper(PCC08),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 )  "
'
'        ' Add By Sindy 98/02/13
'        If Check1.Value = 1 Then
'         'Modify by Amy 2013/12/04 +智權人員/申請國家/總收文號/案件性質/收文日
'         'Modify by Amy 2013/09/30 原:ecd15 AS 狀態
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all SELECT ' ' AS V,ecd02||'-'||LPAD(ecd01,6,'0') AS 編號,NVL(ecd03,'')||NVL(ecd04,'') AS 名稱,NA03 AS 國籍,' ' as 智權人員,'投法開拓'||decode(ecd15,null,null,'-'||ecd15) AS 狀態,ecd16 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM expandcusdetail,expandcusattr,nation Where (instr(NLS_Upper(ecd13),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 ) and ecd10=na01(+) and ecd02=eca01(+) "
'        End If
'        ' 98/02/13 End
'        pub_QL05 = pub_QL05 & ";" & Option1(3).Caption & Trim(txt1(10)) 'Add By Sindy 2010/11/4
'   End If
'
'   If Check1.Value = 1 Then
'      pub_QL05 = pub_QL05 & ";" & Check1.Caption 'Add By Sindy 2010/11/4
'   End If
'
'   CheckOC
'   'modify by nickc 2005/06/03
'   'strSQL = strSQL & " order by 名稱 "
'   '2008/12/3 modify by sonia
'   'strSQL = "select * from (" & strSQL & ") X order by upper(名稱) "
'   'Modify by Amy 2019/09/17 加待活化客戶
'   If Option1(1).Value = True Then
'      'Modify by Amy 2014/01/15 +編號排
'      strSql = "select X.*,Decode(Ocu01,null, '',NVL(Ocu03,0)) as OCU03 from (" & strSql & ") X, OldCustomer Where substr(編號,1,8)= ocu01(+) order by upper(名稱),編號 "
'   Else
'      strSql = "select X.*,Decode(Ocu01,null, '',NVL(Ocu03,0)) as OCU03 from (" & strSql & ") X, OldCustomer Where substr(編號,1,8)= ocu01(+) order by 編號 "
'   End If
'   'end 2019/0917
'   '2008/12/3 end
'   adoRecordset.CursorLocation = adUseClient
'   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'       InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/11/4
'       cmdOK(0).Enabled = True
'       cmdOK(1).Enabled = True
'       cmdOK(2).Enabled = True
'       'add by nickc 2005/10/24
'       cmdOK(5).Enabled = True
'       '911029 nick move from down
'       Set GrdDataList.Recordset = adoRecordset
'   Else
'       InsertQueryLog (0) 'Add By Sindy 2010/11/4
' 'Modify by Amy 2013/12/04 Mark Option1(1).Value = True And Trim(txtfm2(0)) <> "" Then 掉不需再找對造
''       'Add By Sindy 2010/02/05
''       If Option1(1).Value = True And Trim(txtfm2(0)) <> "" Then
''          MsgBox "非本所客戶或代理人，系統會再搜尋案件對造資料，請注意是否有雙方代理情形！", vbInformation
''          Me.Enabled = False
''          If fnSaveParentForm(Me) = False Then
''             Me.Enabled = True
''             Exit Sub
''          End If
''          Screen.MousePointer = vbHourglass
''          frm100110_1.Option1(1).Value = True
''          frm100110_1.txtfm2(0) = Trim(txtfm2(0))
''          frm100110_3.StrMenu
''          Unload frm100110_1
''          Screen.MousePointer = vbDefault
''          Me.Enabled = True
''          Exit Sub
''       '2010/02/05 End
''       Else
'          'Modify by Amy 2013/12/04 +畫面訊息開放可列印
'          Pub_Can_Copy_Pic = True
'          ShowNoData
'          Pub_Can_Copy_Pic = False
'          'end 2013/12/04
'          cmdOK(0).Enabled = False
'          cmdOK(1).Enabled = False
'          cmdOK(2).Enabled = False
'          'add by nickc 2005/10/24
'          cmdOK(5).Enabled = False
'          Screen.MousePointer = vbDefault
'          Exit Sub
''       End If
'   End If
'   Me.GrdDataList.Visible = False 'Add by Amy 2013/12/04
'   CheckOC
'   '911029 nick move to up
'   'Set GrdDataList.Recordset = adoRecordset
'   SetDataListWidth
'
'   'add by nickc 2005/12/14 變色
'   With Me.GrdDataList
'        If .Rows > 0 Then 'Add by Amy 2013/12/04
'            For i = 1 To .Rows - 1
'                .row = i
'                .col = 1
'                .CellForeColor = &H0   '字黑色 'Modfiy by Amy 2019/08/29 原:ForeColor 查儀大會整個變黑
'               'Add by Amy 2019/09/17 活化客戶顯示整列黃,若有呆帳編號底為紅其他欄為黃
'               If .TextMatrix(GrdDataList.row, 15) = "0" And Right(.TextMatrix(GrdDataList.row, 1), 1) <> "＊" Then
'                    For j = 0 To .Cols - 1
'                        '呆帳
'                        If Right(.Text, 1) = "$" And j = 1 Then
'                            .CellBackColor = &HFF& '紅色
'                        '活化客戶
'                        Else
'                            .col = j
'                            .CellBackColor = vbYellow
'                        End If
'                    Next
'                ElseIf Right(.Text, 1) = "$" Then
'                    .CellBackColor = &HFF&
'                'Add by Amy 2019/08/28 +客戶狀態為 遷移不明/解散/廢止/撤銷/停業/死亡 顯示黑底粉字
'                'Modify by Amy 2022/05/23 拿掉 遷移不明 及 停業
'                ElseIf (Left(.Text, 1) = "Y" Or Left(.Text, 1) = "X" Or Left(.Text, 1) = "R") _
'                  And (.TextMatrix(i, 5) = "解散" Or .TextMatrix(i, 5) = "廢止" Or .TextMatrix(i, 5) = "撤銷" Or .TextMatrix(i, 5) = "死亡") Then
'                    For j = 0 To .Cols - 1
'                        .col = j
'                        .CellBackColor = &H0 '黑色
'                        .CellForeColor = &HFF00FF '粉紅色  'Modfiy by Amy 2019/08/29 原:ForeColor
'                    Next j
'                'Add By Sindy 2012/3/21
'                ElseIf Right(.Text, 1) = "⊕" Or .TextMatrix(i, 5) = "對造" Or .TextMatrix(i, 5) = "其他相關人" Then
'                    'Modify by Amy 2013/12/04 對造重抓智權人資料
'                    If Me.GrdDataList.TextMatrix(i, 5) = "對造" Or .TextMatrix(i, 5) = "其他相關人" Then
'                        bolPrint = True '有對造資料
'                        strNo = Pub_RplStr(.TextMatrix(i, 1))
'                        StrToPrint = strNo & ","
'                        Str01 = SystemNumber(strNo, 1)
'                        Select Case Str01
'                            Case "FCP", "FG"
'                                .TextMatrix(i, 4) = GetPrjSalesNM(PUB_GetFCPSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                            Case "FCL", "LIN"
'                                .TextMatrix(i, 4) = GetPrjSalesNM(PUB_GetFCLSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                            Case "FCT"
'                                .TextMatrix(i, 4) = GetPrjSalesNM(PUB_GetFCTSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                            Case "S"
'                                If .TextMatrix(i, 7) = "000" Then
'                                    .TextMatrix(i, 4) = GetPrjSalesNM(PUB_GetFCTSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                                Else
'                                    .TextMatrix(i, 4) = GetPrjSalesNM(PUB_GetAKindSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                                End If
'                            Case Else
'                                .TextMatrix(i, 4) = GetPrjSalesNM(PUB_GetAKindSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                        End Select
'                        .TextMatrix(i, 9) = .TextMatrix(i, 9) & PUB_GetRelateCasePropertyName(.TextMatrix(i, 8), "1")
'                        'Add by Amy 2014/02/25 更新智權人員至暫存檔
'                        strExc(0) = "Update R100102_1 Set R021003='" & .TextMatrix(i, 4) & "' Where R021014='" & Str01 & "' And R021015='" & SystemNumber(strNo, 2) & "' And R021016='" & SystemNumber(strNo, 3) & "' And R021017='" & SystemNumber(strNo, 4) & "' "
'                        cnnConnection.Execute strExc(0)
'                        'end 2014/02/25
'                    End If
'                    'end 2013/12/04
'                    If Right(.Text, 1) = "⊕" Or .TextMatrix(i, 5) = "對造" Then
'                        For j = 0 To .Cols - 1
'                            .col = j
'                            .CellBackColor = &H8080FF
'                        Next j
'                    End If
'                    '2012/3/21 End
'
'                'Add By Sindy 2021/3/25 針對CW03=7.媒介平台，在查詢系統顯示結果為橘色
'                ElseIf Left(.TextMatrix(i, 1), 1) = "平" And .TextMatrix(i, 9) = "7" Then
'                    .CellBackColor = &H80FF& '橘色
'                '2021/3/25 END
'                End If
'
'                'Add by Amy 2020/03/16 國內外潛在客戶 智權人員欄需重抓資料(可能多筆)
'                If Left(.Text, 1) = "R" Then
'                    .TextMatrix(i, 4) = GetDevelopP(.TextMatrix(i, 4))
'                End If
'            Next i
'        End If 'end 2013/12/04
'   End With
'
'   'Add By Cheng 2001/12/26
'   '若查詢結果僅有一筆資料, 則直接勾選
'   If Me.GrdDataList.Rows = 2 Then
'      '911029 nick add
'      GrdDataList.col = 1
'      GrdDataList.row = 1
'      If GrdDataList.Text <> "" Then
'      '911029 nick end
'           GrdDataList.Visible = False
'           GrdDataList.row = 1
'           GrdDataList.col = 0
'           GrdDataList.Text = "V"
'           For i = 0 To GrdDataList.Cols - 1
'               'add by nickc 2005/12/14
'               'Modify By Sindy 2012/3/21 old:If i <> 1 Then
'               If i <> 1 And (i = 2 And Right(GrdDataList.TextMatrix(1, 1), 1) = "⊕") = False Then
'                   GrdDataList.col = i
'                   GrdDataList.CellBackColor = &HFFC0C0
'               End If
'           Next i
'           'Add by Amy 2020/10/15 勾選時判斷有往來記錄,往來記錄鈕變色
'           Call ChkContactRecordBT(GrdDataList.TextMatrix(1, 0), Left(GrdDataList.TextMatrix(1, 1), 8))
'           GrdDataList.Visible = True
'       '911029 nick add
'       End If
'       '911029 nick end
'   End If
'   'Add by Amy 2013/12/04
'   Me.GrdDataList.Visible = True
'   If bolPrint Then
'        cmdOK(8).Enabled = True
'   Else
'        cmdOK(8).Enabled = False
'   End If
'   'end 2013/12/04
'   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
   pub_QL05 = m_pub_QL05 'Add By Sindy 2025/8/13 還原此Form的查詢條件記錄
   If bolFNation = False Then
       s = MsgBox("國內人員不可查詢代理人案件", , "違規.....")
       Unload Me
       Exit Sub
   End If
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   SetDataListWidth
   GetField 'Add by Amy 2023/03/08
   '2011/12/6 modify by sonia
   'txt1(2) = Systemkind_g
   Me.chk.Value = vbChecked
   txt1(2) = "ALL"
   '2011/12/6 end
   'edit by nickc 2007/06/06 切換輸入法改用API
   'txtfm2(0).IMEMode = 1
   Label2(0).Caption = Label2(0).Caption & "／紫底為風險警示" 'Modify by Amy 2024/01/31 +風險檢查對象,拿掉風險警示啟用日
   
   Option1(1).Value = False
   Option2(0).Enabled = False
   Option2(1).Enabled = False
   Option2(2).Enabled = False
   'txtfm2(0).Enabled = False
   Option1(2).Value = False
   'txt1(9).Enabled = False
   cmdOK(0).Enabled = False
   cmdOK(1).Enabled = False
   cmdOK(2).Enabled = False
   'add by nickc 2005/10/24
   cmdOK(5).Enabled = False
   '92.04.16 nick
   cmdState = -1
   
   ' Add By Sindy 98/02/16
   'MODIFY BY SONIA 2015/5/20 因P31及F31人員併入L02,但內外法不開放權限,故改用員工等級控制
   'If Pub_StrUserSt03 = "F31" Or Pub_StrUserSt03 = "F41" Then
   If Pub_strUserST05 >= "51" And Pub_strUserST05 <= "55" Then
      Check1.Value = 1
   Else
      Check1.Value = 0
   End If
   ' 98/02/16 End
   
   'Added by Lydia 2017/12/05 改由啟用日控制
   If strSrvDate(1) >= 國外部關聯企業啟用日 Then cmdOK(2).Caption = "關聯企業" 'Modify by Amy 2023/08/17 原:關聯企業(&R)
   'Add by Amy 2023/08/17 查詢置換字 鈕只有電腦中心才出現
   cmdMemo.Visible = False
   If Pub_StrUserSt03 = "M51" Then cmdMemo.Visible = True
   'end 2023/08/17
   
   m_blnColOrderAsc = True 'Add by Amy 2020/09/04
   
   'Added by Lydia 2018/10/04 來訪通知資料－權限
   m_bExec = IsUserHasRightOfFunction("frm100114_C", strExec, False)
   If m_bExec = False Then
       'Modified by Lydia 2025/06/06
       'CmdWord.Visible = False
       CmdAP(0).Visible = False
       Label10.Visible = False
   End If
   'end 2018/10/04
   
   'add by sonia 2018/10/31 案件統計－業務開拓小組權限
   'modify by sonia 2019/1/2 又增加個人權限
   'm_bExec = IsUserDeptTeam(False)
   m_bExec = IsUserDeptTeam("frm100114_6", False)
   If m_bExec = True Then
     'Modified by Lydia 2018/11/09
      'cmdCase.Visible = True
      cmdOK(9).Visible = True
   Else
      'Modified by Lydia 2018/11/09
      'cmdCase.Visible = False
      cmdOK(9).Visible = False
   End If
   'end 2018/10/04
   
   'Added by Lydia 2025/06/06 互惠期間統計權限同案件統計---Elvan
   If m_bExec = False Then
      CmdAP(1).Visible = False
   Else
      CmdAP(1).Visible = True
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   pub_QL05 = m_pub_QL05 'Add By Sindy 2025/9/12 還原此Form的查詢條件記錄 (多筆查詢有影響)
   Set frm100114_1 = Nothing
End Sub

'Add by Amy 2014/04/25 原寫於grdDataList_SelChange
Private Sub GrdDataList_Click()
   Dim strCopyTxt As String ' Add by Amy 2014/04/25 複製編號文字
    
   grdDataList.row = grdDataList.MouseRow
    
   'Modify by Amy 2014/04/25 +選到編號欄=複製
   grdDataList.col = grdDataList.MouseCol
   If grdDataList.col = 1 Then
        grdDataList.CellForeColor = &H0 '黑色
        'Modify by Amy 2020/09/04 不小心按到欄位名稱也會copy
        If grdDataList.row > 0 Then
            strCopyTxt = grdDataList.TextMatrix(grdDataList.row, grdDataList.col)
        End If
        If strCopyTxt <> "" Then
            '複製編號至剪貼簿
            Clipboard.Clear 'Added by Lydia 2022/01/05 預設清除剪貼簿; 發現Clipboard.SetText前未清除剪貼簿，Ctrl+V貼到Form2.0元件會帶入複製之前的上一筆的複製內容
            Clipboard.SetText strCopyTxt
            grdDataList.CellBackColor = QBColor(7)
            MsgBox "編號已複製", , MsgText(21)
            
            '設回原本顏色
            'Modify by Amy 2023/08/28 改寫至共用函數
'            'Add by Amy 2019/09/17 活化客戶顯示整列黃,若有呆帳編號底為紅其他欄為黃
'            If GrdDataList.TextMatrix(GrdDataList.row, GetValue("待活化客戶")) = "0" And Right(GrdDataList.TextMatrix(GrdDataList.row, GetValue("編號")), 1) <> "＊" Then
'                '呆帳
'                If Right(GrdDataList.Text, 1) = "$" Then
'                    GrdDataList.CellBackColor = &HFF& '紅色
'                '活化客戶
'                Else
'                    GrdDataList.CellBackColor = vbYellow
'                End If
'            'Add by Amy 2019/08/28 +客戶狀態為 遷移不明/解散/廢止/撤銷/停業/死亡 顯示黑底
'            'Modify by Amy 2022/05/23 拿掉 遷移不明 及 停業
'            ElseIf (Left(GrdDataList.Text, 1) = "Y" Or Left(GrdDataList.Text, 1) = "X" Or Left(GrdDataList.Text, 1) = "R") _
'                  And (GrdDataList.TextMatrix(GrdDataList.row, GetValue("狀態")) = "解散" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("狀態")) = "廢止" _
'                   Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("狀態")) = "撤銷" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("狀態")) = "死亡") Then
'                    GrdDataList.CellBackColor = &H0 '黑色
'                    GrdDataList.CellForeColor = &HFF00FF '粉紅色
'            ElseIf Right(GrdDataList.Text, 1) = "⊕" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("狀態")) = "對造" Then
'                GrdDataList.CellBackColor = &H8080FF
'            Else
'                GrdDataList.CellBackColor = QBColor(15)
'            End If
            Call SetMSGridColorQCus(2, Me.Name, grdDataList, strField, IIf(Check3.Value = vbChecked, True, False))
        End If
        Exit Sub
   End If
   'end 2014/04/25
   
   grdDataList.Visible = False
   grdDataList.col = 0
   If grdDataList.row <> 0 Then
        If grdDataList.Text = "V" Then
            grdDataList.Text = ""
            'Modify by Amy 2023/08/28 改寫至共用函數
'            'Add By Sindy 2012/3/21
'            GrdDataList.col = 1
'            'Add by Amy 2019/09/17 活化客戶顯示整列黃,若有呆帳編號底為紅其他欄為黃
'            If GrdDataList.TextMatrix(GrdDataList.row, GetValue("待活化客戶")) = "0" And Right(GrdDataList.TextMatrix(GrdDataList.row, GetValue("編號")), 1) <> "＊" Then
'                For i = 0 To GrdDataList.Cols - 1
'                    '呆帳
'                    If Right(GrdDataList.Text, 1) = "$" And i = 1 Then
'                        GrdDataList.CellBackColor = &HFF& '紅色
'                    '活化客戶
'                    Else
'                        GrdDataList.col = i
'                        GrdDataList.CellBackColor = vbYellow
'                    End If
'                Next
'            'Add by Amy 2019/08/28 +客戶狀態為 遷移不明/解散/廢止/撤銷/停業/死亡 顯示黑底
'            'Modify by Amy 2022/05/23 拿掉 遷移不明 及 停業
'            ElseIf (Left(GrdDataList.Text, 1) = "Y" Or Left(GrdDataList.Text, 1) = "X" Or Left(GrdDataList.Text, 1) = "R") _
'                  And (GrdDataList.TextMatrix(GrdDataList.row, GetValue("狀態")) = "解散" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("狀態")) = "廢止" _
'                   Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("狀態")) = "撤銷" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("狀態")) = "死亡") Then
'                    For i = 0 To GrdDataList.Cols - 1
'                        GrdDataList.col = i
'                        GrdDataList.CellBackColor = &H0 '黑色
'                    Next i
'            ElseIf Right(GrdDataList.Text, 1) = "⊕" Or GrdDataList.TextMatrix(GrdDataList.row, GetValue("狀態")) = "對造" Then
'                For i = 0 To GrdDataList.Cols - 1
'                    GrdDataList.col = i
'                    GrdDataList.CellBackColor = &H8080FF
'                Next i
'            Else
'            '2012/3/21 End
'                For i = 0 To GrdDataList.Cols - 1
'                     'add by nickc 2005/12/14
'                    If i <> 1 Then
'                      GrdDataList.col = i
'                      GrdDataList.CellBackColor = QBColor(15)
'                    End If
'                Next i
'            End If
            Call SetMSGridColorQCus(0, Me.Name, grdDataList, strField, IIf(Check3.Value = vbChecked, True, False))
        '勾選
        Else
            grdDataList.Text = "V"
            'Modify by Amy 2023/08/28 改寫至共用函數
'            For i = 0 To GrdDataList.Cols - 1
'                'add by nickc 2005/12/14
'                'Modify By Sindy 2012/3/21 old:If i <> 1 Then
'                'Modify by Amy 2023/03/08 欄位改動態
'                If i <> 1 And (i = 2 And Right(GrdDataList.TextMatrix(GrdDataList.MouseRow, GetValue("編號")), 1) = "⊕") = False Then
'                   GrdDataList.col = i
'                    GrdDataList.CellBackColor = &HFFC0C0
'                End If
'            Next i
            Call SetMSGridColorQCus(1, Me.Name, grdDataList, strField, IIf(Check3.Value = vbChecked, True, False))
        End If
        'Add by Amy 2020/10/15 勾選時判斷有往來記錄,往來記錄鈕變色
        'Modify by Amy 2023/08/28 bug-聯絡人也會有往來記錄,故拿掉編號只取8碼
        strExc(10) = grdDataList.TextMatrix(grdDataList.row, GetValue("編號"))
        If Left(strExc(10), 1) = "X" Or Left(strExc(10), 1) = "Y" Or Left(strExc(10), 1) = "R" Or Left(strExc(10), 2) = "平台" Then
            Call ChkContactRecordBT(grdDataList.TextMatrix(grdDataList.row, GetValue("V")), strExc(10))
        End If
   End If
   grdDataList.Visible = True
End Sub

'Add by Amy 2020/09/04 +排序
Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If grdDataList.MouseCol < 0 Or grdDataList.MouseRow < 0 Then Exit Sub
    
    grdDataList.col = grdDataList.MouseCol
    grdDataList.row = grdDataList.MouseRow
    If grdDataList.col = 2 Then grdDataList.col = 15 'Modify by Amy 2022/08/19 名稱以OrgN排
    If Me.grdDataList.row < 1 And Me.grdDataList.Text <> "V" Then
        If m_blnColOrderAsc = True Then
            Me.grdDataList.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
        Else
            Me.grdDataList.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
        End If
    End If
End Sub

Private Sub Option1_Click(Index As Integer)
      
   Select Case Index
      Case 0
           If Option1(0).Value = True Then
              'txt1(0).Enabled = True
              Option1(1).Value = False
              Option1(2).Value = False
              Option2(0).Enabled = False
              Option2(1).Enabled = False
              Option2(2).Enabled = False
              Option3(0).Enabled = False
              Option3(1).Enabled = False
              txt1(0).SetFocus
              txt1_GotFocus (0)
              'txtfm2(0).Enabled = False
              'txt1(9).Enabled = False
           End If
      Case 1
           If Option1(1).Value = True Then
              Option2(0).Enabled = True
              Option2(1).Enabled = True
              Option2(2).Enabled = True
              'txtfm2(0).Enabled = True
              
              Option1(0).Value = False
              Option1(2).Value = False
              'txt1(0).Enabled = False
              'txt1(9).Enabled = False
              'txtfm2(0).SetFocus
              'txt1_GotFocus (1)
              Option3(0).Enabled = True
              Option3(1).Enabled = True
              Option3(1).Value = True    '2012/3/28 ADD BY SONIA
              txtFM2(0).SetFocus
           End If
      Case 2
           If Option1(2).Value = True Then
              'txt1(9).Enabled = True
              Option1(0).Value = False
              Option1(1).Value = False
              'txt1(0).Enabled = False
              Option2(0).Enabled = False
              Option2(1).Enabled = False
              Option2(2).Enabled = False
              Option3(0).Enabled = False
              Option3(1).Enabled = False
              'txtfm2(0).Enabled = False
              txt1(9).SetFocus
              txt1_GotFocus (9)
            End If
      'E-MAIL ADD BY Toni 2008/12/03
      Case 3
        If Option1(3).Value = True Then
           Option1(0).Value = False
           Option1(1).Value = False
           Option1(2).Value = False
           
           Option2(0).Enabled = False
           Option2(1).Enabled = False
           Option2(2).Enabled = False
           Option3(0).Enabled = False
           Option3(1).Enabled = False
           
           txt1(10).SetFocus
           txt1_GotFocus (10)
         End If
      Case Else
   End Select
End Sub

Private Sub Option2_Click(Index As Integer)

   Select Case Index
      Case 0
            If Option2(0).Value = True Then
               Option2(1).Value = False
               Option2(2).Value = False
               'Option3(1).Value = True   '2012/3/28 CANCEL BY SONIA 改在Option1(1)
               'edit by nickc 2007/06/06 切換輸入法改用API
               'txtfm2(0).IMEMode = 1
               OpenIme
            End If
      Case 1
            If Option2(1).Value = True Then
               Option2(0).Value = False
               Option2(2).Value = False
               'Option3(1).Value = True   '2012/3/28 CANCEL BY SONIA 改在Option1(1)
               'edit by nickc 2007/06/06 切換輸入法改用API
               'txtfm2(0).IMEMode = 0
               CloseIme
            End If
      Case 2
            If Option2(2).Value = True Then
               Option2(0).Value = False
               Option2(1).Value = False
               'Option3(1).Value = True   '2012/3/28 CANCEL BY SONIA 改在Option1(1)
               'edit by nickc 2007/06/06 切換輸入法改用API
               'txtfm2(0).IMEMode = 1
               OpenIme
            End If
      Case Else
   End Select
   txtFM2(0).SetFocus
   txt1_GotFocus (1)
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   'Mark by Lydia 2022/01/05
'   If Index = 1 Then
'      'If Option2(1).Value = True Then  'Modify by Amy 2013/12/10 改判斷部門
'      If Left(Pub_StrUserSt03, 1) = "F" Then
'         'edit by nickc 2007/06/06 切換輸入法改用API
'         'txtfm2(0).IMEMode = 2
'         CloseIme
'      Else
'         'edit by nickc 2007/06/06 切換輸入法改用API
'         'txtfm2(0).IMEMode = 1
'         OpenIme
'      End If
'   'add by sonia 2014/10/29
'   Else
'      CloseIme
'   'end 2014/10/29
'   End If
   'end 2022/01/05
   
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   '2008/12/3 modify by sonia
   'If Index <> 1 Then
   If Index <> 1 And Index <> 10 Then
   '2008/12/3 end
      KeyAscii = UpperCase(KeyAscii)
   End If
End Sub

Private Sub txt1_LostFocus(Index As Integer)

   Select Case Index
      Case 0
      Case 1
      Case 2 '系統類別
            'Modify By Cheng 2002/03/14
      '      'Add By Cheng 2002/01/07
      '      Me.txt1(Index).Text = GetAllSysKind(Me.txt1(Index))
      Case 3
            If InStr(1, "12", txt1(3)) = 0 Then
                s = MsgBox("查詢別只可 1 或 2 !!", , "USER 輸入錯誤")
                txt1(3).SetFocus
                txt1(3).SelStart = 0
                txt1(3).SelLength = Len(txt1(3))
                Exit Sub
            End If
      Case 4, 5
         If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
            Me.txt1(Index).SetFocus
            txt1_GotFocus Index
            Exit Sub
         End If
         If Index = 5 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
                txt1(Index - 1).SetFocus
                txt1_GotFocus (Index - 1)
                Exit Sub
            End If
          End If
      Case 6
      Case 7
            If RunNick(txt1(Index - 1), txt1(Index)) Then
                txt1(Index - 1).SetFocus
                txt1_GotFocus (Index - 1)
                Exit Sub
            End If
      Case 8
            If Len(txt1(8)) <> 0 Then
                strSql = "SELECT NA03 FROM NATION WHERE NA01='" & txt1(8) & "'"
                CheckOC
                adoRecordset.CursorLocation = adUseClient
                adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
                      If Not IsNull(adoRecordset.Fields(0)) Then
                          LBL1(0).Caption = adoRecordset.Fields(0)
                      Else
                          LBL1(0).Caption = ""
                      End If
                Else
                    LBL1(0).Caption = ""
                    s = MsgBox("國家輸入錯誤！", , "錯誤！")
                    txt1(Index).SetFocus
                    txt1_GotFocus (Index)
                    Exit Sub
                End If
                CheckOC
            Else
                LBL1(0).Caption = ""
            End If
      Case 9
            If Len(txt1(9)) <> 0 Then
                strSql = "SELECT NA03 FROM NATION WHERE NA01='" & txt1(9) & "'"
                CheckOC
                adoRecordset.CursorLocation = adUseClient
                adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
                      If Not IsNull(adoRecordset.Fields(0)) Then
                          LBL1(1).Caption = adoRecordset.Fields(0)
                      Else
                          LBL1(1).Caption = ""
                      End If
                Else
                    LBL1(1).Caption = ""
                    s = MsgBox("國家輸入錯誤！", , "錯誤！")
                    txt1(Index).SetFocus
                    txt1_GotFocus (Index)
                    Exit Sub
                End If
                CheckOC
            Else
                LBL1(1).Caption = ""
            End If
      Case Else
   End Select
End Sub

Private Sub txt1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   
   Select Case Index
      Case 0
          Option1(0).Value = True
      'Mark by Lydia 2022/01/05
      'Case 1
      '    Option1(1).Value = True
      'end 2022/01/05
      Case 9
          Option1(2).Value = True
      Case 10
          Option1(3).Value = True
      Case Else
   End Select
End Sub

'Mark by Amy 2023/09/20 改為共用函數
'Add by Amy 2014/02/25
Private Sub PrintDataA4_Temp()
'    Dim rsPrint As New ADODB.Recordset
'    Dim strPrint As String
'    Dim ii As Integer, jj As Integer
'On Error GoTo Checking
'    intCounter = 1: intRecord = 1: intPage = 1
'
'    Screen.MousePointer = vbHourglass
'    Printer.PaperSize = PUB_GetPaperSize(9) '設定紙張 A4
'    Printer.Orientation = 1 '直印
'    PrintHeadA4
'
'    Printer.FontBold = False
'    'Modify by Amy 2020/09/08 ID+表單
'    strPrint = "Select R021001,R021002,R021003,Decode(R021004,'1','對造','其他相關人'),R021006,R021007,Nvl(To_Char(R021008-19110000),'') " & _
'                 "From R100102_1 Where ID='" & strUserNum & "@" & Me.Name & "' Order by R021002,R021001"
'    intI = 1
'    Set rsPrint = ClsLawReadRstMsg(intI, strPrint)
'    If intI = 1 Then
'        rsPrint.MoveFirst
'        For ii = 0 To rsPrint.RecordCount - 1
'            If intRecord > 45 Then
'                intPage = intPage + 1
'                intRecord = 1
'                Printer.NewPage
'                intCounter = 1
'                PrintHeadA4
'                Printer.FontBold = False
'            End If
'            For jj = 0 To rsPrint.Fields.Count - 1
'                If jj = rsPrint.Fields.Count - 1 Then
'                    Printer.CurrentX = PLeft(jj + 1) - 300 - Printer.TextWidth(rsPrint.Fields(jj).Value) '最右邊
'                Else
'                    Printer.CurrentX = PLeft(jj)
'                End If
'                Printer.CurrentY = 300 + intCounter * 300
'
'                Select Case jj
'                    Case 0 '本所案號
'                        Printer.Print Pub_RplStr(rsPrint.Fields(jj).Value)
'                    Case 1 '名稱
'                        Printer.Print StrToStr(rsPrint.Fields(jj).Value, 10)
'                    Case 2, 3, 4 '智權人員,狀態,總收文號
'                        Printer.Print rsPrint.Fields(jj).Value
'                    Case 5 '案件性質
'                        Printer.Print StrToStr(rsPrint.Fields(jj).Value, 6)
'                    Case 6  '收文日
'                        Printer.Print ChangeTStringToTDateString(rsPrint.Fields(jj).Value)
'                    Case Else
'                End Select
'            Next jj
'            intCounter = intCounter + 1
'            intRecord = intRecord + 1
'            rsPrint.MoveNext
'        Next ii
'    End If
'    Printer.EndDoc
'    Screen.MousePointer = vbDefault
'
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
'   Screen.MousePointer = vbDefault
End Sub
'end 2014/02/25

'Mark by Amy 2014/02/25未使用
'Add by Amy 2013/12/04
Private Sub PrintDataA4()
'    Dim ii As Integer, jj As Integer
'On Error GoTo Checking
'    intCounter = 1: intRecord = 1: intPage = 1
'
'    Screen.MousePointer = vbHourglass
'    Printer.PaperSize = PUB_GetPaperSize(9) '設定紙張 A4
'    Printer.Orientation = 1 '直印
'    PrintHeadA4
'
'    Printer.FontBold = False
'    With Me.GrdDataList
'        For ii = 1 To .Rows - 1
'            If intRecord > 45 Then
'                intPage = intPage + 1
'                intRecord = 1
'                Printer.NewPage
'                intCounter = 1
'                PrintHeadA4
'                Printer.FontBold = False
'            End If
'            'Modify by Amy 2023/03/08 欄位改動態
'            If Left(.TextMatrix(ii, GetValue("編號")), 1) <> "X" And Left(.TextMatrix(ii, GetValue("編號")), 1) <> "Y" And Left(.TextMatrix(ii, GetValue("編號")), 1) <> "R" Then
'                For jj = 1 To .Cols - 1
'                    If jj <= 2 Or jj = 4 Or jj = 5 Or (jj >= 8 And jj <= 10) Then
'                        Select Case jj
'                            Case 1 '本所案號
'                                Printer.CurrentX = PLeft(jj - 1)
'                                Printer.CurrentY = 300 + intCounter * 300
'                                Printer.Print Pub_RplStr(.TextMatrix(ii, jj))
'                            Case 2 '名稱
'                                Printer.CurrentX = PLeft(jj - 1)
'                                Printer.CurrentY = 300 + intCounter * 300
'                                Printer.Print Left(.TextMatrix(ii, jj), 10)
'                            Case 4, 5 '智權人員,狀態
'                                Printer.CurrentX = PLeft(jj - 2)
'                                Printer.CurrentY = 300 + intCounter * 300
'                                Printer.Print .TextMatrix(ii, jj)
'                            Case 8 '總收文號
'                                Printer.CurrentX = PLeft(jj - 4)
'                                Printer.CurrentY = 300 + intCounter * 300
'                                Printer.Print .TextMatrix(ii, jj)
'                            Case 9 '案件性質
'                                Printer.CurrentX = PLeft(jj - 4)
'                                Printer.CurrentY = 300 + intCounter * 300
'                                Printer.Print Left(.TextMatrix(ii, jj), 6)
'                            Case 10  '收文日
'                                Printer.CurrentX = PLeft(jj - 3) - 300 - Printer.TextWidth(.TextMatrix(ii, jj))
'                                Printer.CurrentY = 300 + intCounter * 300
'                                Printer.Print ChangeTStringToTDateString(.TextMatrix(ii, jj))
'                            Case Else
'                        End Select
'                    End If
'                Next jj
'                intCounter = intCounter + 1
'                intRecord = intRecord + 1
'            End If
'        Next ii
'    End With
'    Printer.EndDoc
'    Screen.MousePointer = vbDefault
'
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
'   Screen.MousePointer = vbDefault
End Sub

Private Sub PrintHeadA4()
'   If intPage = 1 Then
'        GetPleft
'        strTp(0) = Me.Caption
'        strTp(1) = ""
'
'        If Option3(0).Value = True Then
'            strTp(1) = strTp(1) & "(字首比對)"
'        ElseIf Option3(1).Value = True Then
'            strTp(1) = strTp(1) & "(模糊比對)"
'        End If
'   End If
'   strTp(2) = "名稱：" & strTp(3) & Space(6) & strTp(1)
'
'   Printer.FontSize = 17
'   Printer.FontBold = True
'   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strTp(0)) / 2)
'   Printer.CurrentY = 300 + intCounter * 300
'   Printer.Print strTp(0)
'
'   Printer.FontSize = 12
'   intCounter = intCounter + 2
'   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strTp(2)) / 2)
'   Printer.CurrentY = 300 + intCounter * 300
'   Printer.Print strTp(2)
'   'Printer.Line (Printer.ScaleWidth / 2 - ((Printer.TextWidth(strTp(2)) - Printer.TextWidth("名稱：")) / 2) + 300, Printer.CurrentY + 30)-(Printer.ScaleWidth / 2 + Printer.TextWidth(strTp(2)) / 2, Printer.CurrentY + 30)
'
'   intCounter = intCounter + 1
'   Printer.CurrentX = 0
'   Printer.CurrentY = 300 + intCounter * 300
'   Printer.Print "操作人員：" & StaffQuery(strUserNum)
'   Printer.CurrentX = 8800
'   Printer.CurrentY = 300 + intCounter * 300
'   Printer.Print "查詢日期：" & CFDate(ACDate(ServerDate))
''   intCounter = intCounter + 1
''   Printer.CurrentX = 12000
''   Printer.CurrentY = 300 + intCounter * 300
''   Printer.Print "頁次: " & intPage
'    intCounter = intCounter + 1
'    For kk = 1 To UBound(PLeft)
'        Printer.CurrentX = PLeft(kk - 1) + (PLeft(kk) - PLeft(kk - 1) - Printer.TextWidth(ColName(kk)) - 100) / 2
'        Printer.CurrentY = 300 + intCounter * 300
'        Printer.Print ColName(kk)
'        Printer.Line (PLeft(kk - 1), Printer.CurrentY)-(PLeft(kk) - 100, Printer.CurrentY)
'    Next kk
'    intCounter = intCounter + 1
End Sub

Private Sub GetPleft()
'   ReDim PLeft(0 To 7)
'   ReDim ColName(1 To 7)
'   PLeft(0) = 100
'   PLeft(1) = PLeft(0) + 2000: ColName(1) = "本所案號"
'   PLeft(2) = PLeft(1) + 2700: ColName(2) = "    名       稱    "
'   PLeft(3) = PLeft(2) + 1200: ColName(3) = "智權人員"
'   PLeft(4) = PLeft(3) + 1500: ColName(4) = " 狀  態 "
'   PLeft(5) = PLeft(4) + 1300: ColName(5) = "總收文號"
'   PLeft(6) = PLeft(5) + 1800: ColName(6) = "案件性質"
'   PLeft(7) = PLeft(6) + 1200: ColName(7) = "收文日"
End Sub
'end 2013/12/04
'end 2023/09/20 不使用

'Added by Lydia 2018/10/04 產生來訪通知資料(Word)
'Modified by Lydia 2025/06/06 改成陣列
'Private Sub cmdWord_Click()
Private Sub CmdAP_Click(Index As Integer)
Dim strTmp1 As String

   Me.Enabled = False
   strTmp1 = ""

   For i = 1 To grdDataList.Rows - 1
        grdDataList.col = 0
        grdDataList.row = i
        If Trim(grdDataList.Text) = "V" Then
           grdDataList.col = 0
           grdDataList.Text = ""
           grdDataList.col = 1
           strTmp1 = Pub_RplStr(grdDataList.Text)
           Exit For
        End If
   Next i
   
   Me.Enabled = True
   'Added by Lydia 2025/06/06
   Select Case Index
      Case 0 '產生來訪通知資料(Word)
   'end 2025/06/06
         If strTmp1 = "" Then
            MsgBox "請選擇代理人編號 !", vbCritical
            txt1(0).SetFocus
            txt1_GotFocus 0
            Exit Sub
         End If
         'Added by Lydia 2025/09/19
JumpToReInput2:
         'Modified by Lydia 2025/11/11 改選項說明
         'm_strTotKind = InputBox("請輸入統計方式：1-新申請案    2-案件數" & vbCrLf & "空白=取消", "來訪資料(Word)", "1")
         m_strTotKind = InputBox("請輸入統計方式：1-新案（委任申請案）" & vbCrLf & "  2-在案（目前代理案）　　空白=取消", "來訪資料(Word)", "1")
         If m_strTotKind = "" Then
            Exit Sub
         Else
            If m_strTotKind <> "1" And m_strTotKind <> "2" Then
               GoTo JumpToReInput2
            End If
         End If
         'end 2025/09/19
   
         '取得系統日之年度
         If m_YY(0) = "" Then
            m_YY(0) = PUB_DBYEAR(strSrvDate(1))
            m_YY(1) = PUB_DBYEAR(CompDate(0, -1, strSrvDate(1)))
            m_YY(2) = PUB_DBYEAR(CompDate(0, -2, strSrvDate(1)))
            m_YY(3) = PUB_DBYEAR(CompDate(0, -3, strSrvDate(1)))
            m_YY(4) = PUB_DBYEAR(CompDate(0, -4, strSrvDate(1)))
            m_YY(5) = PUB_DBYEAR(CompDate(0, -5, strSrvDate(1)))
         End If
   
         '產生Word
         Erase m_Item
         iUpper = 0
         Screen.MousePointer = vbHourglass
         'Modified by Lydia 2025/06/06
         'cmdWord.Enabled = False
         CmdAP(0).Enabled = False
         'Modified by Lydia 2018/11/06 改寫法
         Call DoWordNew(ChangeCustomerL(Left(strTmp1, 8)))
         'Modified by Lydia 2025/06/06
         'cmdWord.Enabled = True
         CmdAP(0).Enabled = True
         Screen.MousePointer = vbDefault
         Erase m_Item
   'Added by Lydia 2025/06/06
      Case 1
         If strTmp1 = "" Then
            MsgBox "請選擇代理人編號 !", vbCritical
            txt1(0).SetFocus
            txt1_GotFocus 0
            Exit Sub
         End If
         
         Call frm050408.SetParent(Me, strTmp1)
         Me.Hide
         frm050408.Show
   End Select
   'end 2025/06/06
End Sub

'Modified by Lydia 2018/11/06 改寫法
Private Sub DoWordNew(ByVal pNo As String)
Dim rsRD As New ADODB.Recordset
Dim strGrp As String
Dim strDateYear As String
Dim strPA As String, strTM As String, strLCall As String, strSP As String
Dim strCon1 As String
Dim strConPA As String, strConTM As String, strConLC As String, strConSP As String
Dim strTemp(0 To 2) As String
Dim intR As Integer
Dim iCall As Integer
Dim intJ As Integer
Dim tmpArr As Variant, tmpArr2 As Variant
Dim mRgrng  'As Range 'Remove by Lydia 2018/11/06 取消型態(Casher會出錯)
Dim mESeqNo As String '暫存檔序號
Dim strFAlist As String '所有關係企業編號
Dim mRg1 As Integer, mRg2 As Integer '合併表格的起始和終止
Dim bVisible As Boolean 'Added by Lydia 2019/04/09
Dim strMidCon As String 'Added by Lydia 2019/05/06
Dim intCR As Integer 'Added by Lydia 2025/08/22 往來記錄筆數
Dim strNewPA As String, strNewTM As String, strNewLC As String, strNewSP As String 'Added by Lydia 2025/09/19

    bolRetry = False

On Error GoTo ErrHandle
    
    'Added by Lydia 2025/08/22 清除查詢印表記錄檔欄位
    ClearQueryLog (Me.Name)
    'Modified by Lydia 2025/09/19
    'Modified by Lydia 2025/11/11 改選項說明:1-新申請案    2-案件數=>1-新案（委任申請案）    2-在案（目前代理案）
    pub_QL05 = pub_QL05 & ";" & pNo & "(來訪資料Word);統計方式:" & IIf(m_strTotKind = "1", "1-新案（委任申請案）", "2-在案（目前代理案）")
    'end 2025/08/22
    'Added by Lydia 2019/05/06 清除暫存檔
    cnnConnection.Execute "DELETE FROM R100114_6 WHERE ID='" & strUserNum & "' AND FORMID = '" & Me.Name & "' "
    
    '抓所有關係企業
    If Left(pNo, 1) = "Y" Then
          strSql = "select '1' as ord1, fa01||fa02 as fno, Decode(fa05,null,nvl(fa04,fa06),fa05||' '||fa63||' '||fa64||' '||fa65) as fname,na03 " & _
                      "from fagent,nation where fa01='" & Left(pNo, 8) & "' and fa02='0' and fa10=na01(+) " & _
                      "union all select '2' as ord1, fa01||fa02 as fno, Decode(fa05,null,nvl(fa04,fa06),fa05||' '||fa63||' '||fa64||' '||fa65) as fname,na03 " & _
                      "from fagent,nation where substr(fa01,1,6)='" & Left(pNo, 6) & "' and fa01 <> '" & Left(pNo, 8) & "' and fa02='0' and fa10=na01(+) "
    ElseIf Left(pNo, 1) = "X" Then
          strSql = "select '1' as ord1,cu01||cu02 as fno, Decode(cu05,null,nvl(cu04,cu06),cu05||' '||cu88||' '||cu89||' '||cu90) as fname,na03 " & _
                      "from customer,nation where cu01='" & Left(pNo, 8) & "' and cu02='0' and cu10=na01(+) " & _
                      "union all select '2' as ord1,cu01||cu02 as fno, Decode(cu05,null,nvl(cu04,cu06),cu05||' '||cu88||' '||cu89||' '||cu90) as fname,na03 " & _
                      "from customer,nation where substr(cu01,1,6)='" & Left(pNo, 6) & "' and cu01 <> '" & Left(pNo, 8) & "' and cu02='0' and cu10=na01(+) "
    'Added by Lydia 2023/01/11 潛在客戶
    Else
          strSql = "select '1' as ord1,pcu01||pcu02 as fno, Decode(pcu03,null,nvl(pcu08,pcu07),pcu03||' '||pcu04||' '||pcu05||' '||pcu06) as fname,na03 " & _
                      "from potcustomer,nation where pcu01='" & Left(pNo, 8) & "' and pcu02='0' and pcu09=na01(+) " & _
                      "union all select '1' as ord1,poc01||poc02 as fno, Decode(poc23,null,nvl(poc03,poc27),poc23||' '||poc24||' '||poc25||' '||poc26) as fname,na03 " & _
                      "from potcustomer1,nation where poc01='" & Left(pNo, 8) & "'  and poc02='0' and poc04=na01(+) "
    'end 2023/01/11
    End If
    strSql = strSql & " order by ord1, fno asc "
    intR = 0
    Set rsRD = ClsLawReadRstMsg(intR, strSql)
    If intR = 1 Then
          rsRD.MoveFirst
          strTemp(0) = Trim("" & rsRD.Fields("fname")) '主要的代理人/客戶名稱
          strTemp(1) = Trim("" & rsRD.Fields("na03")) '國籍
          'Move by Lydia 2019/05/06 移到上面 (取得系統別和剔除特定案件性質)
          'Modified by Lydia 2025/09/17 剔除特定案件性質改成與客戶/代理人案件統計的「案件往來」一致
           'For intR = 0 To 10
           '     strExc(intR) = ""
           'Next intR
           'strExc(5) = GetCaseClosePtyList(1, strExc(1), True)
           'strExc(6) = GetCaseClosePtyList(2, strExc(2), True)
           'strExc(7) = GetCaseClosePtyList(5, strExc(3), True)
           'strExc(8) = GetCaseClosePtyList(3, strExc(4), True)
           ''Added by Lydia 2019/05/06
           'strMidCon = "AND CP44 IS NOT NULL AND CP158>19221111 AND CP09<'C' AND CP01||CP10 NOT IN (" & Replace(strExc(5) & "," & strExc(6) & "," & strExc(7) & "," & strExc(8), ",,", ",") & ")" 'Added by Lydia 2019/05/06 CF案件判斷案件日期為最小發文日;並且針對1.現在案件屬於A, 2.最初案件屬於A, 3.中間案件屬於A的情況都要抓到,所以先將資料丟暫存檔
           strExc(1) = SQLGrpStr("", 1) '專利
           strExc(2) = SQLGrpStr("", 2) '商標
           strExc(3) = SQLGrpStr("", 5) '服務
           strExc(4) = SQLGrpStr("", 3)  '法務
           strMidCon = "AND CP44 IS NOT NULL AND CP158>19221111 AND CP09<'C' AND CP01||CP10 NOT IN (" & cntFAnotCP10tot & ") "
           'end 2025/09/19
           
           'Added by Lydia 2025/09/19
           If m_strTotKind = "1" Then '選擇統計方式：1-新申請案
              strNewPA = " and (cp01 in (" & Replace(strExc(1), ",' '", "") & ") and " & PUB_GetForNewCaseSql("1") & ") "
              strNewTM = " and (cp01 in (" & Replace(strExc(2), ",' '", "") & ") " & PUB_GetForNewCaseSql("2") & ") "
              strNewSP = " and (cp01 in (" & Replace(strExc(3), ",' '", "") & ") " & PUB_GetForNewCaseSql("5") & ") "
              strNewLC = " and cp01 in (" & Replace(strExc(4), ",' '", "") & ") "
           End If
           'end 2025/09/19
           
          'Added by Lydia 2023/01/11 潛在客戶
          If Left("" & rsRD.Fields("fno"), 1) = "R" Then
              '所有關係企業: 用；區隔不同編號, 用|區隔國籍和名稱
              strFAlist = strFAlist & "；" & rsRD.Fields("fno") & "|" & Trim("" & rsRD.Fields("na03")) & Trim("" & rsRD.Fields("fname")) & "(" & rsRD.Fields("fno") & ")"
          Else
          'end 2023/01/13
             '逐筆處理案件統計
             Do While Not rsRD.EOF
                   If Left("" & rsRD.Fields("fno"), 1) = "Y" Then
                          strConPA = " pa75 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%")
                          strConTM = " tm44 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%")
                          strConSP = " sp26 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%")
                          strConLC = " lc27 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%")
                   ElseIf Left("" & rsRD.Fields("fno"), 1) = "X" Then
                         strConPA = " (pa26 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%") & " or pa27 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%") & " or pa28 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%") & " or pa29 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%") & " or pa30 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%") & ") "
                         strConTM = " (tm23 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%") & " or tm78 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%") & " or tm79 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%") & " or tm80 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%") & " or tm81 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%") & ") "
                         strConSP = " (sp08 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%") & " or sp58 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%") & " or sp59 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%") & " or sp65 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%") & " or sp66 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%") & ") "
                         strConLC = " (lc11 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%") & " or lc43 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%") & " or lc44 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%") & " or lc45 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%") & " or lc46 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%") & ") "
                   End If
                   'Added by Lydia 2019/05/06 CF案件判斷案件日期為最小發文日;並且針對1.現在案件屬於A, 2.最初案件屬於A, 3.中間案件屬於A的情況都要抓到,所以先將資料丟暫存檔
                   strSql = "INSERT INTO R100114_6 (ID,FORMID,PNO,C01,C02,C03,C04,MINCP09,MAXCP09) " & _
                                "SELECT '" & strUserNum & "','" & Me.Name & "','" & Left("" & rsRD.Fields("fno"), 8) & "' ,CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09,MAX(CP27||CP09||CP44) MAXCP09 FROM CASEPROGRESS " & _
                                "   WHERE (CP01,CP02,CP03,CP04) IN (SELECT DISTINCT CP01,CP02,CP03,CP04 FROM CASEPROGRESS,PATENT,TRADEMARK WHERE CP158>19221111 " & _
                                "   AND CP158<>NVL(NVL(TM30,PA58),0) AND CP44 LIKE '" & Left("" & rsRD.Fields("fno"), 8) & "%' AND CP04='00' AND CP09<'C' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & _
                                "   AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP44||' '<>PA75||TM44||' ') " & strMidCon & _
                                "   GROUP BY CP01,CP02,CP03,CP04"
                   cnnConnection.Execute strSql, intI
         
                    '專利(非CF代理人)
                    'P大陸案(PA75)=>FMP ; CFP案(PA75)=>FCFP
                    'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
                    'strPA = "SELECT '10' as ord1, " & IIf(Left("" & rsRd.Fields("fno"), 1) = "Y", "DECODE(CP01,'CFP','FCFP','P',DECODE(PA09,'000',CP01,'FMP'),CP01)", "") & " cp01,cp02,cp03,cp04,decode(pa08,'1','發明','2','新型','3','設計',pa08) skind,pa08,1 as cnt,cp05 as mdate1 " & _
                                 "from caseprogress,patent where cp09 in (select substr(min(cp05||cp09),9) from caseprogress,patent " & _
                                          "where " & strConPA & " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp09 < 'C' and cp05 > 19221111 and (cp158 > 19221111 or (cp158=0 and cp159=0)) group by cp01,cp02,cp03,cp04) " & _
                                 "and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa01 is not null"
                    'Modified by Lydia 2025/09/19 + strNewPA
                    strPA = "SELECT '10' as ord1, " & IIf(Left("" & rsRD.Fields("fno"), 1) = "Y", "DECODE(CP01,'CFP','FCFP','P',DECODE(PA09,'000',CP01,'FMP'),CP01)", "") & " cp01,cp02,cp03,cp04,decode(pa08,'1','發明','2','新型','3','設計',pa08) skind,pa08,1 as cnt,cp05 as mdate1 " & _
                                 "from caseprogress,patent where cp09 in (select substr(min(cp05||cp09),9) from caseprogress,patent " & _
                                          "where " & strConPA & strNewPA & " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp09 < 'C' and cp05 > 19221111 and ((CP158>19221111 AND CP158<>NVL(PA58,0)) OR (CP158=0 AND CP159=0)) group by cp01,cp02,cp03,cp04) " & _
                                 "and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa01 is not null"
                            
                  '專利(代理人)
                  If Left("" & rsRD.Fields("fno"), 1) = "Y" Then
                       'CF代理人範圍
                     'Added by Lydia 2025/09/19
                     If m_strTotKind = "1" Then '選擇統計方式：1-新申請案，不用另外分析案件已轉它所。
                       strPA = strPA & " union SELECT '11' as ord1, cp01,cp02,cp03,cp04,decode(pa08,'1','發明','2','新型','3','設計',pa08) skind,pa08,1 as cnt,cp27 as mdate1 " & _
                                  "from caseprogress,patent where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) " & _
                                  "and cp09 IN (SELECT SUBSTR(MINCP09,9,9) FROM " & _
                                  "(SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09 FROM CASEPROGRESS WHERE (CP01,CP02,CP03,CP04) IN " & _
                                  "(SELECT C01,C02,C03,C04 FROM R100114_6 WHERE ID='" & strUserNum & "' AND FORMID='" & Me.Name & "' AND INSTR(MINCP09,'" & Left("" & rsRD.Fields("fno"), 8) & "') > 0 AND PNO='" & Left("" & rsRD.Fields("fno"), 8) & "' ) " & _
                                  "AND CP44 LIKE '" & Left("" & rsRD.Fields("fno"), 8) & "%' " & strMidCon & strNewPA & _
                                  " GROUP BY CP01,CP02,CP03,CP04)) AND PA01 IS NOT NULL"
                     Else
                     'end 2025/09/19
                     
                      'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
                       'strCon1 = "SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09,MAX(CP27||CP09||CP44) MAXCP09 FROM CASEPROGRESS " & _
                                       "WHERE (CP01,CP02,CP03,CP04) IN (SELECT DISTINCT CP01,CP02,CP03,CP04 FROM CASEPROGRESS,PATENT WHERE CP158>19221111 AND CP44 LIKE " & CNULL(Left("" & rsRd.Fields("fno"), 8) & "%") & " AND CP04='00' AND CP09<'C' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA01 IS NOT NULL AND CP44||' '<>PA75||' ') " & _
                                          "AND CP158>19221111 AND CP09<'C' " & IIf(strExc(5) <> "", "AND CP01||CP10 NOT IN (" & strExc(5) & ")", "") & " GROUP BY CP01,CP02,CP03,CP04 "
                        'Modified by Lydia 2019/05/06 現在案件屬於Y編號
                        'strCon1 = "SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09,MAX(CP27||CP09||CP44) MAXCP09 FROM CASEPROGRESS " & _
                                        "WHERE (CP01,CP02,CP03,CP04) IN (SELECT DISTINCT CP01,CP02,CP03,CP04 FROM CASEPROGRESS,PATENT WHERE CP158>19221111 AND CP158<>NVL(PA58,0) AND CP44 LIKE " & CNULL(Left("" & rsRd.Fields("fno"), 8) & "%") & " AND CP04='00' AND CP09<'C' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA01 IS NOT NULL AND CP44||' '<>PA75||' ') " & _
                                        "AND CP158>19221111 AND CP09<'C' " & IIf(strExc(5) <> "", "AND CP01||CP10 NOT IN (" & strExc(5) & ")", "") & " GROUP BY CP01,CP02,CP03,CP04 "
                        'strPA = strPA & " union SELECT '11' as ord1, cp01,cp02,cp03,cp04,decode(pa08,'1','發明','2','新型','3','設計',pa08) skind,pa08,1 as cnt,cp27 as mdate1 " & _
                                   "from caseprogress,patent where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) " & _
                                   "and cp09 in (SELECT SUBSTR(MAXCP09,9,9) FROM (" & strCon1 & ")) AND CP44 LIKE " & CNULL(Left("" & rsRd.Fields("fno"), 8) & "%")
                       strPA = strPA & " union SELECT '11' as ord1, cp01,cp02,cp03,cp04,decode(pa08,'1','發明','2','新型','3','設計',pa08) skind,pa08,1 as cnt,cp27 as mdate1 " & _
                                  "from caseprogress,patent where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) " & _
                                  "and cp09 IN (SELECT SUBSTR(MINCP09,9,9) FROM " & _
                                  "(SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09 FROM CASEPROGRESS WHERE (CP01,CP02,CP03,CP04) IN " & _
                                  "(SELECT C01,C02,C03,C04 FROM R100114_6 WHERE ID='" & strUserNum & "' AND FORMID='" & Me.Name & "' AND INSTR(MAXCP09,'" & Left("" & rsRD.Fields("fno"), 8) & "') > 0 AND PNO='" & Left("" & rsRD.Fields("fno"), 8) & "' ) " & _
                                  "AND CP44 LIKE '" & Left("" & rsRD.Fields("fno"), 8) & "%' " & strMidCon & _
                                  " GROUP BY CP01,CP02,CP03,CP04)) AND PA01 IS NOT NULL"
                               
                        '更換代理人(加註*)
                       'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
                       'strPA = strPA & " union SELECT '12' as ord1,DECODE(CP01,'CFP','FCFP','P',DECODE(PA09,'000',CP01,'FMP'),CP01)||'*' cp01,cp02,cp03,cp04,decode(pa08,'1','發明','2','新型','3','設計',pa08) skind,pa08,1 as cnt,cp05 as mdate1 " & _
                                  "from caseprogress,patent where cp09 in (select substr(min(cp05||cp09),9) from caseprogress,patent " & _
                                             "where cp139 like " & CNULL(Left("" & rsRd.Fields("fno"), 8) & "%") & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) " & _
                                             "and pa01 is not null and cp09 < 'C' and substr(cp139,1,8)<>substr(pa75,1,8) and cp05 > 19221111 and (cp158 > 19221111 or (cp158=0 and cp159=0)) group by cp01,cp02,cp03,cp04) " & _
                                   "and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
                       '更換FC代理人
                       strPA = strPA & " union SELECT '12' as ord1,DECODE(CP01,'CFP','FCFP','P',DECODE(PA09,'000',CP01,'FMP'),CP01)||'*' cp01,cp02,cp03,cp04,decode(pa08,'1','發明','2','新型','3','設計',pa08) skind,pa08,1 as cnt,cp05 as mdate1 " & _
                                  "from caseprogress,patent where cp09 in (select substr(min(cp05||cp09),9) from caseprogress,patent " & _
                                             "where cp139 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%") & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) " & _
                                             "and pa01 is not null and cp09 < 'C' and substr(cp139,1,8)<>substr(pa75,1,8) and cp05 > 19221111 and ((CP158>19221111 AND CP158<>NVL(PA58,0)) OR (CP158=0 AND CP159=0)) group by cp01,cp02,cp03,cp04) " & _
                                   "and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
                       '更換CF代理人
                       'Modified by Lydia 2019/05/06 現在案件不屬於Y編號
                       'strPA = strPA & " union SELECT '13' as ord1, cp01||'*' as cp01,cp02,cp03,cp04,decode(pa08,'1','發明','2','新型','3','設計',pa08) skind,pa08,1 as cnt,cp27 as mdate1 " & _
                                  "from caseprogress,patent where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) " & _
                                  "and cp09 in (SELECT SUBSTR(MINCP09,9,9) FROM (" & strCon1 & ") WHERE SUBSTR(MINCP09,18)<>SUBSTR(MAXCP09,18)) AND CP44 LIKE " & CNULL(Left("" & rsRd.Fields("fno"), 8) & "%")
                       strPA = strPA & " union SELECT '13' as ord1, cp01||'*' as cp01,cp02,cp03,cp04,decode(pa08,'1','發明','2','新型','3','設計',pa08) skind,pa08,1 as cnt,cp27 as mdate1 " & _
                                  "from caseprogress,patent where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) " & _
                                  "and cp09 IN (SELECT SUBSTR(MINCP09,9,9) FROM " & _
                                  "(SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09 FROM CASEPROGRESS WHERE (CP01,CP02,CP03,CP04) IN " & _
                                  "(SELECT C01,C02,C03,C04 FROM R100114_6 WHERE ID='" & strUserNum & "' AND FORMID='" & Me.Name & "' AND INSTR(MAXCP09,'" & Left("" & rsRD.Fields("fno"), 8) & "') = 0 AND PNO='" & Left("" & rsRD.Fields("fno"), 8) & "' ) " & _
                                  "AND CP44 LIKE '" & Left("" & rsRD.Fields("fno"), 8) & "%' " & strMidCon & _
                                  " GROUP BY CP01,CP02,CP03,CP04)) AND PA01 IS NOT NULL"
                     End If 'Added by Lydia 2025/09/19
                  End If
               
                  '商標(非CF代理人)
                  'T大陸案(TM44)=>FMT ; CFT案(TM44)=>FCFT
                  'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
                   'strTM = "SELECT '20' as ord1, " & IIf(Left("" & rsRd.Fields("fno"), 1) = "Y", "DECODE(CP01,'CFT','FCFT','T',DECODE(TM10,'000',CP01,'FMT'),CP01)", "") & " cp01,cp02,cp03,cp04,'申請' skind,'1',1 as cnt,cp05 as mdate1 " & _
                               "from caseprogress,TRADEMARK where cp09 in (select substr(min(cp05||cp09),9) from caseprogress,TRADEMARK " & _
                                        "where " & strConTM & " and cp01(+)=TM01 and cp02(+)=TM02 and cp03(+)=TM03 and cp04(+)=TM04 and cp09 < 'C' and cp05 > 19221111 and (cp158 > 19221111 or (cp158=0 and cp159=0)) group by cp01,cp02,cp03,cp04) " & _
                               "and cp01=TM01(+) and cp02=TM02(+) and cp03=TM03(+) and cp04=TM04(+) and TM01 is not null"
                  'Modified by Lydia 2020/12/09 為了避免誤解案件類型,統一拿掉"申請" ; '申請' skind, => null skind
                  'Modified by Lydia 2025/09/19 + strNewTM
                  strTM = "SELECT '20' as ord1, " & IIf(Left("" & rsRD.Fields("fno"), 1) = "Y", "DECODE(CP01,'CFT','FCFT','T',DECODE(TM10,'000',CP01,'FMT'),CP01)", "") & " cp01,cp02,cp03,cp04,null skind,'1',1 as cnt,cp05 as mdate1 " & _
                               "from caseprogress,TRADEMARK where cp09 in (select substr(min(cp05||cp09),9) from caseprogress,TRADEMARK " & _
                                        "where " & strConTM & strNewTM & " and cp01(+)=TM01 and cp02(+)=TM02 and cp03(+)=TM03 and cp04(+)=TM04 and cp09 < 'C' and cp05 > 19221111 and ((CP158>19221111 AND CP158<>NVL(TM30,0)) OR (CP158=0 AND CP159=0)) group by cp01,cp02,cp03,cp04) " & _
                               "and cp01=TM01(+) and cp02=TM02(+) and cp03=TM03(+) and cp04=TM04(+) and TM01 is not null"
                  '商標(代理人)
                  If Left("" & rsRD.Fields("fno"), 1) = "Y" Then
                       'CF代理人範圍
                     'Added by Lydia 2025/09/19
                     If m_strTotKind = "1" Then '選擇統計方式：1-新申請案，不用另外分析案件已轉它所。
                       strTM = strTM & " union SELECT '21' as ord1, cp01,cp02,cp03,cp04,null skind,'1',1 as cnt,cp27 as mdate1 " & _
                                  "from caseprogress,TRADEMARK where cp01=TM01(+) and cp02=TM02(+) and cp03=TM03(+) and cp04=TM04(+) " & _
                                  "and cp09 IN (SELECT SUBSTR(MINCP09,9,9) FROM " & _
                                  "(SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09 FROM CASEPROGRESS WHERE (CP01,CP02,CP03,CP04) IN " & _
                                  "(SELECT C01,C02,C03,C04 FROM R100114_6 WHERE ID='" & strUserNum & "' AND FORMID='" & Me.Name & "' AND INSTR(MINCP09,'" & Left("" & rsRD.Fields("fno"), 8) & "') > 0 AND PNO='" & Left("" & rsRD.Fields("fno"), 8) & "' ) " & _
                                  "AND CP44 LIKE '" & Left("" & rsRD.Fields("fno"), 8) & "%' " & strMidCon & strNewTM & _
                                  " GROUP BY CP01,CP02,CP03,CP04)) AND TM01 IS NOT NULL"
                     Else
                     'end 2025/09/19

                       'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
                       'strCon1 = "SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09,MAX(CP27||CP09||CP44) MAXCP09 FROM CASEPROGRESS " & _
                                       "WHERE (CP01,CP02,CP03,CP04) IN (SELECT DISTINCT CP01,CP02,CP03,CP04 FROM CASEPROGRESS,TRADEMARK WHERE CP158>19221111 AND CP44 LIKE " & CNULL(Left("" & rsRd.Fields("fno"), 8) & "%") & " AND CP04='00' AND CP09<'C' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM01 IS NOT NULL AND CP44||' '<>TM44||' ') " & _
                                       "AND CP158>19221111 AND CP09<'C' " & IIf(strExc(6) <> "", "AND CP01||CP10 NOT IN (" & strExc(6) & ")", "") & " GROUP BY CP01,CP02,CP03,CP04 "
                       'Modified by Lydia 2019/05/06 現在案件屬於Y編號
                       'strCon1 = "SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09,MAX(CP27||CP09||CP44) MAXCP09 FROM CASEPROGRESS " & _
                                       "WHERE (CP01,CP02,CP03,CP04) IN (SELECT DISTINCT CP01,CP02,CP03,CP04 FROM CASEPROGRESS,TRADEMARK WHERE CP158>19221111 AND CP158<>NVL(TM30,0) AND CP44 LIKE " & CNULL(Left("" & rsRd.Fields("fno"), 8) & "%") & " AND CP04='00' AND CP09<'C' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM01 IS NOT NULL AND CP44||' '<>TM44||' ') " & _
                                       "AND CP158>19221111 AND CP09<'C' " & IIf(strExc(6) <> "", "AND CP01||CP10 NOT IN (" & strExc(6) & ")", "") & " GROUP BY CP01,CP02,CP03,CP04 "
                       'strTM = strTM & " union SELECT '21' as ord1, cp01,cp02,cp03,cp04,'申請' skind,'1',1 as cnt,cp27 as mdate1 " & _
                                  "from caseprogress,TRADEMARK where cp01=TM01(+) and cp02=TM02(+) and cp03=TM03(+) and cp04=TM04(+) " & _
                                  "and cp09 in (SELECT SUBSTR(MAXCP09,9,9) FROM (" & strCon1 & ")) AND CP44 LIKE " & CNULL(Left("" & rsRd.Fields("fno"), 8) & "%")
                       'Modified by Lydia 2020/12/09 為了避免誤解案件類型,統一拿掉"申請" ; '申請' skind, => null skind
                       strTM = strTM & " union SELECT '21' as ord1, cp01,cp02,cp03,cp04,null skind,'1',1 as cnt,cp27 as mdate1 " & _
                                  "from caseprogress,TRADEMARK where cp01=TM01(+) and cp02=TM02(+) and cp03=TM03(+) and cp04=TM04(+) " & _
                                  "and cp09 IN (SELECT SUBSTR(MINCP09,9,9) FROM " & _
                                  "(SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09 FROM CASEPROGRESS WHERE (CP01,CP02,CP03,CP04) IN " & _
                                  "(SELECT C01,C02,C03,C04 FROM R100114_6 WHERE ID='" & strUserNum & "' AND FORMID='" & Me.Name & "' AND INSTR(MAXCP09,'" & Left("" & rsRD.Fields("fno"), 8) & "') > 0 AND PNO='" & Left("" & rsRD.Fields("fno"), 8) & "' ) " & _
                                  "AND CP44 LIKE '" & Left("" & rsRD.Fields("fno"), 8) & "%' " & strMidCon & _
                                  " GROUP BY CP01,CP02,CP03,CP04)) AND TM01 IS NOT NULL"
                               
                       '更換FC代理人(加註*)
                       'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
                       'strTM = strTM & " union SELECT '22' as ord1,DECODE(CP01,'CFT','FCFT','T',DECODE(TM10,'000',CP01,'FMT'),CP01)||'*' cp01,cp02,cp03,cp04,'申請' skind,'1',1 as cnt,cp05 as mdate1 " & _
                                  "from caseprogress,TRADEMARK where cp09 in (select substr(min(cp05||cp09),9) from caseprogress,TRADEMARK " & _
                                             "where cp139 like " & CNULL(Left("" & rsRd.Fields("fno"), 8) & "%") & " and cp01=TM01(+) and cp02=TM02(+) and cp03=TM03(+) and cp04=TM04(+) " & _
                                             "and TM01 is not null and cp09 < 'C' and substr(cp139,1,8)<>substr(TM44,1,8) and cp05 > 19221111 and (cp158 > 19221111 or (cp158=0 and cp159=0)) group by cp01,cp02,cp03,cp04) " & _
                                   "and cp01=TM01(+) and cp02=TM02(+) and cp03=TM03(+) and cp04=TM04(+) "
                       'Modified by Lydia 2020/12/09 為了避免誤解案件類型,統一拿掉"申請" ; '申請' skind, => null skind
                       strTM = strTM & " union SELECT '22' as ord1,DECODE(CP01,'CFT','FCFT','T',DECODE(TM10,'000',CP01,'FMT'),CP01)||'*' cp01,cp02,cp03,cp04,null skind,'1',1 as cnt,cp05 as mdate1 " & _
                                  "from caseprogress,TRADEMARK where cp09 in (select substr(min(cp05||cp09),9) from caseprogress,TRADEMARK " & _
                                             "where cp139 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%") & " and cp01=TM01(+) and cp02=TM02(+) and cp03=TM03(+) and cp04=TM04(+) " & _
                                             "and TM01 is not null and cp09 < 'C' and substr(cp139,1,8)<>substr(TM44,1,8) and cp05 > 19221111 and ((CP158>19221111 AND CP158<>NVL(TM30,0)) OR (CP158=0 AND CP159=0)) group by cp01,cp02,cp03,cp04) " & _
                                   "and cp01=TM01(+) and cp02=TM02(+) and cp03=TM03(+) and cp04=TM04(+) "
                       '更換CF代理人
                       'Modified by Lydia 2019/05/06 現在案件不屬於Y編號
                       'strTM = strTM & " union SELECT '23' as ord1, cp01||'*' as cp01,cp02,cp03,cp04,'申請' skind,'1',1 as cnt,cp27 as mdate1 " & _
                                  "from caseprogress,TRADEMARK where cp01=TM01(+) and cp02=TM02(+) and cp03=TM03(+) and cp04=TM04(+) " & _
                                  "and cp09 in (SELECT SUBSTR(MINCP09,9,9) FROM (" & strCon1 & ") WHERE SUBSTR(MINCP09,18)<>SUBSTR(MAXCP09,18)) AND CP44 LIKE " & CNULL(Left("" & rsRd.Fields("fno"), 8) & "%")
                       'Modified by Lydia 2020/12/09 為了避免誤解案件類型,統一拿掉"申請" ; '申請' skind, => null skind
                       strTM = strTM & " union SELECT '23' as ord1, cp01||'*' as cp01,cp02,cp03,cp04,null skind,'1',1 as cnt,cp27 as mdate1 " & _
                                  "from caseprogress,TRADEMARK where cp01=TM01(+) and cp02=TM02(+) and cp03=TM03(+) and cp04=TM04(+) " & _
                                  "and cp09 IN (SELECT SUBSTR(MINCP09,9,9) FROM " & _
                                  "(SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09 FROM CASEPROGRESS WHERE (CP01,CP02,CP03,CP04) IN " & _
                                  "(SELECT C01,C02,C03,C04 FROM R100114_6 WHERE ID='" & strUserNum & "' AND FORMID='" & Me.Name & "' AND INSTR(MAXCP09,'" & Left("" & rsRD.Fields("fno"), 8) & "') = 0 AND PNO='" & Left("" & rsRD.Fields("fno"), 8) & "' ) " & _
                                  "AND CP44 LIKE '" & Left("" & rsRD.Fields("fno"), 8) & "%' " & strMidCon & _
                                  " GROUP BY CP01,CP02,CP03,CP04)) AND TM01 IS NOT NULL"
                     End If 'Added by Lydia 2025/09/19
                  End If

                  '服務(非CF代理人)
                  'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
                   'strSP = "SELECT '30' as ord1, cp01,cp02,cp03,cp04,null skind,'1',1 as cnt,cp05 as mdate1 " & _
                             "from caseprogress,SERVICEPRACTICE where cp09 in (select substr(min(cp05||cp09),9) from caseprogress,SERVICEPRACTICE " & _
                                        "where " & strConSP & " and cp01(+)=SP01 and cp02(+)=SP02 and cp03(+)=SP03 and cp04(+)=SP04 and cp09 < 'C' and cp05 > 19221111 and (cp158 > 19221111 or (cp158=0 and cp159=0)) group by cp01,cp02,cp03,cp04) " & _
                               "and cp01=SP01(+) and cp02=SP02(+) and cp03=SP03(+) and cp04=SP04(+) and SP01 is not null"
                   'Modified by Lydia 2025/09/19 +strNewSP
                   strSP = "SELECT '30' as ord1, cp01,cp02,cp03,cp04,null skind,'1',1 as cnt,cp05 as mdate1 " & _
                             "from caseprogress,SERVICEPRACTICE where cp09 in (select substr(min(cp05||cp09),9) from caseprogress,SERVICEPRACTICE " & _
                                        "where " & strConSP & strNewSP & " and cp01(+)=SP01 and cp02(+)=SP02 and cp03(+)=SP03 and cp04(+)=SP04 and cp09 < 'C' and cp05 > 19221111 and ((CP158>19221111 AND CP158<>NVL(SP16,0)) OR (CP158=0 AND CP159=0)) group by cp01,cp02,cp03,cp04) " & _
                               "and cp01=SP01(+) and cp02=SP02(+) and cp03=SP03(+) and cp04=SP04(+) and SP01 is not null"
                  '服務(代理人)
                  If Left("" & rsRD.Fields("fno"), 1) = "Y" Then
                    'CF代理人範圍
                     'Added by Lydia 2025/09/19
                     If m_strTotKind = "1" Then '選擇統計方式：1-新申請案，不用另外分析案件已轉它所。
                       strSP = strSP & " union SELECT '31' as ord1, cp01,cp02,cp03,cp04,null skind,'1',1 as cnt,cp27 as mdate1 " & _
                                  "from caseprogress,SERVICEPRACTICE where cp01=SP01(+) and cp02=SP02(+) and cp03=SP03(+) and cp04=SP04(+) " & _
                                  "and cp09 IN (SELECT SUBSTR(MINCP09,9,9) FROM " & _
                                  "(SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09 FROM CASEPROGRESS WHERE (CP01,CP02,CP03,CP04) IN " & _
                                  "(SELECT C01,C02,C03,C04 FROM R100114_6 WHERE ID='" & strUserNum & "' AND FORMID='" & Me.Name & "' AND INSTR(MINCP09,'" & Left("" & rsRD.Fields("fno"), 8) & "') > 0 AND PNO='" & Left("" & rsRD.Fields("fno"), 8) & "' ) " & _
                                  "AND CP44 LIKE '" & Left("" & rsRD.Fields("fno"), 8) & "%' " & strMidCon & strNewSP & _
                                  " GROUP BY CP01,CP02,CP03,CP04)) AND SP01 IS NOT NULL"
                     Else
                     'end 2025/09/19
                       'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
                       'strCon1 = "SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09,MAX(CP27||CP09||CP44) MAXCP09 FROM CASEPROGRESS " & _
                                       "WHERE (CP01,CP02,CP03,CP04) IN (SELECT DISTINCT CP01,CP02,CP03,CP04 FROM CASEPROGRESS,SERVICEPRACTICE WHERE CP158>19221111 AND CP44 LIKE " & CNULL(Left("" & rsRd.Fields("fno"), 8) & "%") & " AND CP04='00' AND CP09<'C' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP01 IS NOT NULL AND CP44||' '<>SP26||' ') " & _
                                       "AND CP158>19221111 AND CP09<'C' " & IIf(strExc(7) <> "", "AND CP01||CP10 NOT IN (" & strExc(7) & ")", "") & " GROUP BY CP01,CP02,CP03,CP04 "
                       'Modified by Lydia 2019/05/06 現在案件屬於Y編號
                       'strCon1 = "SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09,MAX(CP27||CP09||CP44) MAXCP09 FROM CASEPROGRESS " & _
                                       "WHERE (CP01,CP02,CP03,CP04) IN (SELECT DISTINCT CP01,CP02,CP03,CP04 FROM CASEPROGRESS,SERVICEPRACTICE WHERE CP158>19221111 AND CP158<>NVL(SP16,0) AND CP44 LIKE " & CNULL(Left("" & rsRd.Fields("fno"), 8) & "%") & " AND CP04='00' AND CP09<'C' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP01 IS NOT NULL AND CP44||' '<>SP26||' ') " & _
                                       "AND CP158>19221111 AND CP09<'C' " & IIf(strExc(7) <> "", "AND CP01||CP10 NOT IN (" & strExc(7) & ")", "") & " GROUP BY CP01,CP02,CP03,CP04 "
                       'strSP = strSP & " union SELECT '31' as ord1, cp01,cp02,cp03,cp04,null skind,'1',1 as cnt,cp27 as mdate1 " & _
                                  "from caseprogress,SERVICEPRACTICE where cp01=SP01(+) and cp02=SP02(+) and cp03=SP03(+) and cp04=SP04(+) " & _
                                  "and cp09 in (SELECT SUBSTR(MAXCP09,9,9) FROM (" & strCon1 & ")) AND CP44 LIKE " & CNULL(Left("" & rsRd.Fields("fno"), 8) & "%")
                       strSP = strSP & " union SELECT '31' as ord1, cp01,cp02,cp03,cp04,null skind,'1',1 as cnt,cp27 as mdate1 " & _
                                  "from caseprogress,SERVICEPRACTICE where cp01=SP01(+) and cp02=SP02(+) and cp03=SP03(+) and cp04=SP04(+) " & _
                                  "and cp09 IN (SELECT SUBSTR(MINCP09,9,9) FROM " & _
                                  "(SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09 FROM CASEPROGRESS WHERE (CP01,CP02,CP03,CP04) IN " & _
                                  "(SELECT C01,C02,C03,C04 FROM R100114_6 WHERE ID='" & strUserNum & "' AND FORMID='" & Me.Name & "' AND INSTR(MAXCP09,'" & Left("" & rsRD.Fields("fno"), 8) & "') > 0 AND PNO='" & Left("" & rsRD.Fields("fno"), 8) & "' ) " & _
                                  "AND CP44 LIKE '" & Left("" & rsRD.Fields("fno"), 8) & "%' " & strMidCon & _
                                  " GROUP BY CP01,CP02,CP03,CP04)) AND SP01 IS NOT NULL"
                    
                       '更換FC代理人(加註*)
                       'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
                       'strSP = strSP & " union SELECT '32' as ord1,cp01||'*' cp01,cp02,cp03,cp04,null skind,'1',1 as cnt,cp05 as mdate1 " & _
                                  "from caseprogress,SERVICEPRACTICE where cp09 in (select substr(min(cp05||cp09),9) from caseprogress,SERVICEPRACTICE " & _
                                             "where cp139 like " & CNULL(Left("" & rsRd.Fields("fno"), 8) & "%") & " and cp01=SP01(+) and cp02=SP02(+) and cp03=SP03(+) and cp04=SP04(+) " & _
                                             "and SP01 is not null and cp09 < 'C' and substr(cp139,1,8)<>substr(SP26,1,8) and cp05 > 19221111 and (cp158 > 19221111 or (cp158=0 and cp159=0)) group by cp01,cp02,cp03,cp04) " & _
                                   "and cp01=SP01(+) and cp02=SP02(+) and cp03=SP03(+) and cp04=SP04(+) "
                       strSP = strSP & " union SELECT '32' as ord1,cp01||'*' cp01,cp02,cp03,cp04,null skind,'1',1 as cnt,cp05 as mdate1 " & _
                                  "from caseprogress,SERVICEPRACTICE where cp09 in (select substr(min(cp05||cp09),9) from caseprogress,SERVICEPRACTICE " & _
                                             "where cp139 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%") & " and cp01=SP01(+) and cp02=SP02(+) and cp03=SP03(+) and cp04=SP04(+) " & _
                                             "and SP01 is not null and cp09 < 'C' and substr(cp139,1,8)<>substr(SP26,1,8) and cp05 > 19221111 and ((CP158>19221111 AND CP158<>NVL(SP16,0)) OR (CP158=0 AND CP159=0)) group by cp01,cp02,cp03,cp04) " & _
                                   "and cp01=SP01(+) and cp02=SP02(+) and cp03=SP03(+) and cp04=SP04(+) "
                       '更換CF代理人
                       'Modified by Lydia 2019/05/06 現在案件屬於Y編號
                       'strSP = strSP & " union SELECT '33' as ord1, cp01||'*' as cp01,cp02,cp03,cp04,null skind,'1',1 as cnt,cp27 as mdate1 " & _
                                  "from caseprogress,SERVICEPRACTICE where cp01=SP01(+) and cp02=SP02(+) and cp03=SP03(+) and cp04=SP04(+) " & _
                                  "and cp09 in (SELECT SUBSTR(MINCP09,9,9) FROM (" & strCon1 & ") WHERE SUBSTR(MINCP09,18)<>SUBSTR(MAXCP09,18)) AND CP44 LIKE " & CNULL(Left("" & rsRd.Fields("fno"), 8) & "%")
                       strSP = strSP & " union SELECT '33' as ord1, cp01||'*' as cp01,cp02,cp03,cp04,null skind,'1',1 as cnt,cp27 as mdate1 " & _
                                  "from caseprogress,SERVICEPRACTICE where cp01=SP01(+) and cp02=SP02(+) and cp03=SP03(+) and cp04=SP04(+) " & _
                                  "and cp09 IN (SELECT SUBSTR(MINCP09,9,9) FROM " & _
                                  "(SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09 FROM CASEPROGRESS WHERE (CP01,CP02,CP03,CP04) IN " & _
                                  "(SELECT C01,C02,C03,C04 FROM R100114_6 WHERE ID='" & strUserNum & "' AND FORMID='" & Me.Name & "' AND INSTR(MAXCP09,'" & Left("" & rsRD.Fields("fno"), 8) & "') = 0 AND PNO='" & Left("" & rsRD.Fields("fno"), 8) & "' ) " & _
                                  "AND CP44 LIKE '" & Left("" & rsRD.Fields("fno"), 8) & "%' " & strMidCon & _
                                  " GROUP BY CP01,CP02,CP03,CP04)) AND SP01 IS NOT NULL"
                     End If 'Added by Lydia 2025/09/19
                  End If
               
                  '法務(非CF代理人)
                  'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
                   'strLCall = "SELECT '40' as ord1, " & IIf(Left("" & rsRd.Fields("fno"), 1) = "Y", "DECODE(CP01,'CFL','FCFL','L',DECODE(LC15,'000',CP01,'FML'),CP01)", "") & " cp01,cp02,cp03,cp04,null skind,'1',1 as cnt,cp05 as mdate1 " & _
                             "from caseprogress,LAWCASE where cp09 in (select substr(min(cp05||cp09),9) from caseprogress,LAWCASE " & _
                                        "where " & strConLC & " and cp01(+)=LC01 and cp02(+)=LC02 and cp03(+)=LC03 and cp04(+)=LC04 and cp09 < 'C' and cp05 > 19221111 and (cp158 > 19221111 or (cp158=0 and cp159=0)) group by cp01,cp02,cp03,cp04) " & _
                               "and cp01=LC01(+) and cp02=LC02(+) and cp03=LC03(+) and cp04=LC04(+) and LC01 is not null"
                   'Modified by Lydia 2025/09/19 + strNewLC
                   strLCall = "SELECT '40' as ord1, " & IIf(Left("" & rsRD.Fields("fno"), 1) = "Y", "DECODE(CP01,'CFL','FCFL','L',DECODE(LC15,'000',CP01,'FML'),CP01)", "") & " cp01,cp02,cp03,cp04,null skind,'1',1 as cnt,cp05 as mdate1 " & _
                             "from caseprogress,LAWCASE where cp09 in (select substr(min(cp05||cp09),9) from caseprogress,LAWCASE " & _
                                        "where " & strConLC & strNewLC & " and cp01(+)=LC01 and cp02(+)=LC02 and cp03(+)=LC03 and cp04(+)=LC04 and cp09 < 'C' and cp05 > 19221111 and ((CP158>19221111 AND CP158<>NVL(LC09,0)) OR (CP158=0 AND CP159=0)) group by cp01,cp02,cp03,cp04) " & _
                               "and cp01=LC01(+) and cp02=LC02(+) and cp03=LC03(+) and cp04=LC04(+) and LC01 is not null"
                  '法務(代理人)
                  If Left("" & rsRD.Fields("fno"), 1) = "Y" Then
                       'CF代理人範圍
                     'Added by Lydia 2025/09/19
                     If m_strTotKind = "1" Then '選擇統計方式：1-新申請案，不用另外分析案件已轉它所。
                       strLCall = strLCall & " union SELECT '41' as ord1, cp01,cp02,cp03,cp04,null skind,'1',1 as cnt,cp27 as mdate1 " & _
                                  "from caseprogress,LAWCASE where cp01=LC01(+) and cp02=LC02(+) and cp03=LC03(+) and cp04=LC04(+) " & _
                                  "and cp09 IN (SELECT SUBSTR(MINCP09,9,9) FROM " & _
                                  "(SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09 FROM CASEPROGRESS WHERE (CP01,CP02,CP03,CP04) IN " & _
                                  "(SELECT C01,C02,C03,C04 FROM R100114_6 WHERE ID='" & strUserNum & "' AND FORMID='" & Me.Name & "' AND INSTR(MINCP09,'" & Left("" & rsRD.Fields("fno"), 8) & "') > 0 AND PNO='" & Left("" & rsRD.Fields("fno"), 8) & "' ) " & _
                                  "AND CP44 LIKE '" & Left("" & rsRD.Fields("fno"), 8) & "%' " & strMidCon & strNewLC & _
                                  " GROUP BY CP01,CP02,CP03,CP04)) AND LC01 IS NOT NULL"
                     Else
                     'end 2025/09/19
                       'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
                       'strCon1 = "SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09,MAX(CP27||CP09||CP44) MAXCP09 FROM CASEPROGRESS " & _
                                       "WHERE (CP01,CP02,CP03,CP04) IN (SELECT DISTINCT CP01,CP02,CP03,CP04 FROM CASEPROGRESS,LAWCASE WHERE CP158>19221111 AND CP44 LIKE " & CNULL(Left("" & rsRd.Fields("fno"), 8) & "%") & " AND CP04='00' AND CP09<'C' AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND LC01 IS NOT NULL AND CP44||' '<>LC27||' ') " & _
                                       "AND CP158>19221111 AND CP09<'C' " & IIf(strExc(8) <> "", "AND CP01||CP10 NOT IN (" & strExc(8) & ")", "") & " GROUP BY CP01,CP02,CP03,CP04 "
                       'Modified by Lydia 2019/05/06 現在案件屬於Y編號
                       'strCon1 = "SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09,MAX(CP27||CP09||CP44) MAXCP09 FROM CASEPROGRESS " & _
                                       "WHERE (CP01,CP02,CP03,CP04) IN (SELECT DISTINCT CP01,CP02,CP03,CP04 FROM CASEPROGRESS,LAWCASE WHERE CP158>19221111 AND CP158<>NVL(LC09,0) AND CP44 LIKE " & CNULL(Left("" & rsRd.Fields("fno"), 8) & "%") & " AND CP04='00' AND CP09<'C' AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND LC01 IS NOT NULL AND CP44||' '<>LC27||' ') " & _
                                       "AND CP158>19221111 AND CP09<'C' " & IIf(strExc(8) <> "", "AND CP01||CP10 NOT IN (" & strExc(8) & ")", "") & " GROUP BY CP01,CP02,CP03,CP04 "
                       'strLCall = strLCall & " union SELECT '41' as ord1, cp01,cp02,cp03,cp04,null skind,'1',1 as cnt,cp27 as mdate1 " & _
                                  "from caseprogress,LAWCASE where cp01=LC01(+) and cp02=LC02(+) and cp03=LC03(+) and cp04=LC04(+) " & _
                                  "and cp09 in (SELECT SUBSTR(MAXCP09,9,9) FROM (" & strCon1 & ")) AND CP44 LIKE " & CNULL(Left("" & rsRd.Fields("fno"), 8) & "%")
                       strLCall = strLCall & " union SELECT '41' as ord1, cp01,cp02,cp03,cp04,null skind,'1',1 as cnt,cp27 as mdate1 " & _
                                  "from caseprogress,LAWCASE where cp01=LC01(+) and cp02=LC02(+) and cp03=LC03(+) and cp04=LC04(+) " & _
                                  "and cp09 IN (SELECT SUBSTR(MINCP09,9,9) FROM " & _
                                  "(SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09 FROM CASEPROGRESS WHERE (CP01,CP02,CP03,CP04) IN " & _
                                  "(SELECT C01,C02,C03,C04 FROM R100114_6 WHERE ID='" & strUserNum & "' AND FORMID='" & Me.Name & "' AND INSTR(MAXCP09,'" & Left("" & rsRD.Fields("fno"), 8) & "') > 0 AND PNO='" & Left("" & rsRD.Fields("fno"), 8) & "' ) " & _
                                  "AND CP44 LIKE '" & Left("" & rsRD.Fields("fno"), 8) & "%' " & strMidCon & _
                                  " GROUP BY CP01,CP02,CP03,CP04)) AND LC01 IS NOT NULL"
                       '更換FC代理人(加註*)
                       'Modified by Lydia 2019/01/02 P-120097,P-120098都是發文當天就閉卷，應該不能列入計算件數。
                       'strLCall = strLCall & " union SELECT '42' as ord1,DECODE(CP01,'CFL','FCFL','L',DECODE(LC15,'000',CP01,'FML'),CP01)||'*' cp01,cp02,cp03,cp04,null skind,'1',1 as cnt,cp05 as mdate1 " & _
                                  "from caseprogress,LAWCASE where cp09 in (select substr(min(cp05||cp09),9) from caseprogress,LAWCASE " & _
                                             "where cp139 like " & CNULL(Left("" & rsRd.Fields("fno"), 8) & "%") & " and cp01=LC01(+) and cp02=LC02(+) and cp03=LC03(+) and cp04=LC04(+) " & _
                                             "and LC01 is not null and cp09 < 'C' and substr(cp139,1,8)<>substr(LC27,1,8) and cp05 > 19221111 and (cp158 > 19221111 or (cp158=0 and cp159=0)) group by cp01,cp02,cp03,cp04) " & _
                                   "and cp01=LC01(+) and cp02=LC02(+) and cp03=LC03(+) and cp04=LC04(+) "
                       strLCall = strLCall & " union SELECT '42' as ord1,DECODE(CP01,'CFL','FCFL','L',DECODE(LC15,'000',CP01,'FML'),CP01)||'*' cp01,cp02,cp03,cp04,null skind,'1',1 as cnt,cp05 as mdate1 " & _
                                  "from caseprogress,LAWCASE where cp09 in (select substr(min(cp05||cp09),9) from caseprogress,LAWCASE " & _
                                             "where cp139 like " & CNULL(Left("" & rsRD.Fields("fno"), 8) & "%") & " and cp01=LC01(+) and cp02=LC02(+) and cp03=LC03(+) and cp04=LC04(+) " & _
                                             "and LC01 is not null and cp09 < 'C' and substr(cp139,1,8)<>substr(LC27,1,8) and cp05 > 19221111 and ((CP158>19221111 AND CP158<>NVL(LC09,0)) OR (CP158=0 AND CP159=0)) group by cp01,cp02,cp03,cp04) " & _
                                   "and cp01=LC01(+) and cp02=LC02(+) and cp03=LC03(+) and cp04=LC04(+) "
                       '更換CF代理人
                       'strLCall = strLCall & " union SELECT '43' as ord1, cp01||'*' as cp01,cp02,cp03,cp04,null skind,'1',1 as cnt,cp27 as mdate1 " & _
                                  "from caseprogress,LAWCASE where cp01=LC01(+) and cp02=LC02(+) and cp03=LC03(+) and cp04=LC04(+) " & _
                                  "and cp09 in (SELECT SUBSTR(MINCP09,9,9) FROM (" & strCon1 & ") WHERE SUBSTR(MINCP09,18)<>SUBSTR(MAXCP09,18)) AND CP44 LIKE " & CNULL(Left("" & rsRd.Fields("fno"), 8) & "%")
                       strLCall = strLCall & " union SELECT '43' as ord1, cp01||'*' as cp01,cp02,cp03,cp04,null skind,'1',1 as cnt,cp27 as mdate1 " & _
                                  "from caseprogress,LAWCASE where cp01=LC01(+) and cp02=LC02(+) and cp03=LC03(+) and cp04=LC04(+) " & _
                                  "and cp09 IN (SELECT SUBSTR(MINCP09,9,9) FROM " & _
                                  "(SELECT CP01,CP02,CP03,CP04,MIN(CP27||CP09||CP44) MINCP09 FROM CASEPROGRESS WHERE (CP01,CP02,CP03,CP04) IN " & _
                                  "(SELECT C01,C02,C03,C04 FROM R100114_6 WHERE ID='" & strUserNum & "' AND FORMID='" & Me.Name & "' AND INSTR(MAXCP09,'" & Left("" & rsRD.Fields("fno"), 8) & "') = 0 AND PNO='" & Left("" & rsRD.Fields("fno"), 8) & "' ) " & _
                                  "AND CP44 LIKE '" & Left("" & rsRD.Fields("fno"), 8) & "%' " & strMidCon & _
                                  " GROUP BY CP01,CP02,CP03,CP04)) AND LC01 IS NOT NULL"
                     End If 'Added by Lydia 2025/09/19
                  End If

                   '歷年統計
                   strSql = strPA & " Union " & strTM & " Union " & strSP & " Union " & strLCall
                   'Modifed by Lydia 2019/01/02 改排序:專利1->商標2->服務3->法務->4
                   'strSql = "select substr(mdate1,1,4) yyyy,ord1,cp01,pa08,skind,sum(cnt) totcnt from (" & strSql & ") group by substr(mdate1,1,4) ,ord1,cp01,pa08,skind order by ord1,cp01,pa08,1 "
                   strSql = "select substr(mdate1,1,4) yyyy,substr(ord1,1,1) as ord1,cp01,pa08,skind,sum(cnt) totcnt from (" & strSql & ") group by substr(mdate1,1,4) ,substr(ord1,1,1),cp01,pa08,skind order by 2 asc,cp01 asc,pa08 asc,1 "
                
                   intR = 1
                   Set RsTemp = ClsLawReadRstMsg(intR, strSql)
                   strGrp = ""
                   If intR = 1 Then
                        RsTemp.MoveFirst
                        Do While Not RsTemp.EOF
                            If strGrp <> Trim("" & RsTemp.Fields("cp01") & RsTemp.Fields("pa08")) Then '系統別+種類
                                iUpper = iUpper + 1
                                ReDim Preserve m_Item(1 To iUpper + 1) '宣告陣列
                                '代理人資料
                                SetItemWordArray m_Item, iUpper, 1, Trim("" & rsRD.Fields("na03")) & Trim("" & rsRD.Fields("fname")) & "(" & rsRD.Fields("fno") & ")"
                                '排序
                                SetItemWordArray m_Item, iUpper, 2, Format(iUpper, "00")
                                '系統別
                                SetItemWordArray m_Item, iUpper, 3, Trim("" & RsTemp.Fields("cp01")) & Trim("" & RsTemp.Fields("skind"))
                                strGrp = Trim("" & RsTemp.Fields("cp01") & RsTemp.Fields("pa08"))
                            End If
                            intJ = 0
                            If Trim("" & RsTemp.Fields("yyyy")) = m_YY(0) Then '當年
                                 intJ = 9
                            ElseIf InStr(m_YY(1) & "," & m_YY(2) & "," & m_YY(3) & "," & m_YY(4) & "," & m_YY(5), "" & RsTemp.Fields("yyyy")) > 0 Then   '前1~5年
                                 intJ = 9 - (Val(m_YY(0)) - Val("" & RsTemp.Fields("yyyy")))
                            End If

                            If intJ > 0 Then '年度統計
                                   SetItemWordArray m_Item, iUpper, intJ, "" & RsTemp.Fields("totcnt")
                            End If
                            '歷年統計
                            SetItemWordArray m_Item, iUpper, 10, "" & RsTemp.Fields("totcnt")
                         
                            RsTemp.MoveNext
                        Loop '新案案件數
                   End If
                 
                   '所有關係企業: 用；區隔不同編號, 用|區隔國籍和名稱
                   strFAlist = strFAlist & "；" & rsRD.Fields("fno") & "|" & Trim("" & rsRD.Fields("na03")) & Trim("" & rsRD.Fields("fname")) & "(" & rsRD.Fields("fno") & ")"
                                      
                   rsRD.MoveNext
             Loop  '----關係企業
          End If ' If Left("" & rsRd.Fields("fno"), 1) <> "R" Then 'Added by Lydia 2023/01/11 潛在客戶
          
          strFAlist = Mid(strFAlist, 2)
    End If '新案案件數: 統計完成

'---------------------------------------
JumpToWord:
    If strFAlist <> "" Then
          '開啟Word檔
          'Modifiec by Lydia 2019/04/09 改成模組
'          If TypeName(g_WordAp) <> "Application" Then Set g_WordAp = New Word.Application
'
'             '判斷word是否已開啟
'          If g_WordAp Is Nothing Then
'RestarWord:
'              Set g_WordAp = New Word.Application
'              g_WordAp.Visible = False
'          End If
'
'          g_WordAp.Visible = True 'Added by Lydia 2019/04/09 更換office2013後,需要顯示Word畫面才可正常產生檔案
'          g_WordAp.Documents.add
RestarWord:
      If Pub_NewWordDoc(g_WordAp, bVisible, m_WordLeft, m_WordTop) = True Then
'end 2019/04/09
          With g_WordAp.Application
                '版面設定
                .Selection.PageSetup.Orientation = wdOrientPortrait '直印
                .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
                .Selection.PageSetup.RightMargin = .CentimetersToPoints(2)
                .Selection.PageSetup.TopMargin = .CentimetersToPoints(2)
                .Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
                .Selection.PageSetup.FooterDistance = .CentimetersToPoints(2)
                .Selection.Orientation = wdTextOrientationHorizontal
                .Selection.Font.Name = "標楷體"
                
                .Selection.Font.Size = 14
                .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Selection.TypeText "國外" & IIf(Left(pNo, 1) = "Y", "代理人", "廠商") & "來訪通知"
                .Selection.TypeParagraph
                .Selection.Font.Size = 12
      
                '新增表格(1*N)
                .Selection.Tables.add Range:=.Selection.Range, NumRows:=1, NumColumns:=2
                '畫格線
                With .Selection.Tables(1)
                    .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                    .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                    .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
                    .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
                    .Borders(wdBorderHorizontal).LineStyle = wdLineStyleSingle
                    .Borders(wdBorderVertical).LineStyle = wdLineStyleSingle
                    .Borders.Shadow = False
                End With
                .Selection.SelectRow
                .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
               .Selection.Paragraphs.Alignment = wdAlignParagraphLeft
                '.Selection.Rows.SpaceBetweenColumns = CentimetersToPoints(0.1) '容易出錯
                .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(3), RulerStyle:=wdAdjustProportional
                .Selection.Collapse Direction:=wdCollapseStart
                .Selection.Paragraphs.Alignment = wdAlignParagraphCenter
                '----------
                intJ = 1

                .Selection.TypeText Text:=IIf(Left(pNo, 1) = "Y", "代理人", "廠商") & "名稱"
                .Selection.MoveRight Unit:=wdCell, Count:=1
                .Selection.TypeText Text:="***" & strTemp(1) & " " & strTemp(0) & "(" & pNo & ")***"
                .Selection.MoveRight Unit:=wdCell, Count:=1
                '----------
                intJ = intJ + 1
                .Selection.Paragraphs.Alignment = wdAlignParagraphCenter
                .Selection.TypeText Text:="來訪人員"
                .Selection.MoveRight Unit:=wdCell, Count:=2
                '---------
                intJ = intJ + 1
                .Selection.Paragraphs.Alignment = wdAlignParagraphCenter
                .Selection.TypeText Text:="來訪時間"
                .Selection.MoveRight Unit:=wdCell, Count:=2
                '---------
                intJ = intJ + 1
                .Selection.Paragraphs.Alignment = wdAlignParagraphCenter
                .Selection.TypeText Text:="接待人員"
                .Selection.MoveRight Unit:=wdCell, Count:=2
                '---------
                If iUpper > 0 Then '有案件統計
                    intJ = intJ + 1
                    .Selection.Paragraphs.Alignment = wdAlignParagraphCenter
                    .Selection.TypeText Text:="案件往來"
                    .Selection.MoveRight Unit:=wdCell, Count:=1
                    'Added by Lydia 2025/09/19
                    If m_strTotKind = "1" Then
                        'Modified by Lydia 2025/11/11 加註統計方式
                        .Selection.TypeText Text:="案件僅統計新申請案性質的案件數。統計方式：1-新案（委任申請案）"
                    Else
                    'end 2025/09/19
                        'Modified by Lydia 2025/11/11 加註統計方式
                        .Selection.TypeText Text:="案件如已更代或申請人已變動，則以系統類別*呈現此種狀況中其為原代理人或原申請人的案件數。統計方式：2-在案（目前代理案）" '案件往來備註
                    End If 'Added by Lydia 2025/09/19
                    .Selection.MoveRight Unit:=wdCell, Count:=2
                    intJ = intJ + 1
                    .Selection.Paragraphs.Alignment = wdAlignParagraphCenter
                    .Selection.Cells.Split NumRows:=1, NumColumns:=8, MergeBeforeSplit:=False
                    .Selection.SelectRow
                    .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.3), RulerStyle:=wdAdjustProportional
                    .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
                    .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
                    .Selection.Cells(5).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
                    .Selection.Cells(6).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
                    .Selection.Cells(7).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
                    .Selection.Cells(8).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
                    '改成關係企業的案件統計量
                    .Selection.SelectRow
                    .Selection.Collapse Direction:=wdCollapseStart
                    .Selection.MoveRight Unit:=wdCell, Count:=1
                    '抬頭年度
                    For iCall = 5 To 0 Step -1
                         .Selection.MoveRight Unit:=wdCell, Count:=1
                         .Selection.TypeText Text:=m_YY(iCall)
                    Next iCall
                    .Selection.Font.Shading.BackgroundPatternColorIndex = wdNoHighlight
                    .Selection.MoveRight Unit:=wdCell, Count:=1
                    .Selection.TypeText Text:="歷年合計"
                    .Selection.MoveRight Unit:=wdCell, Count:=1 '跳下一行
                    intJ = intJ + 1
                    strGrp = ""
                    '讀取案件統計陣列
                    For iCall = 1 To iUpper
                         '代理人資料
                         If strGrp <> m_Item(iCall).IA01 Then
                             If InStr(m_Item(iCall).IA01, pNo) > 0 Then
                                 .Selection.TypeText Text:="***" & m_Item(iCall).IA01 & "***"
                             Else
                                 .Selection.TypeText Text:=m_Item(iCall).IA01
                             End If
                             If strGrp <> "" Then '不同代理人,要先合併前代理人的欄位
                                 mRg2 = intJ - 1
                                 If mRg2 > mRg1 Then
                                     Set mRgrng = .Selection.Tables(1).Cell(mRg1, 1).Range
                                    mRgrng.End = .Selection.Tables(1).Cell(mRg2, 1).Range.End
                                    mRgrng.Select
                                    .Selection.Cells.Merge '合併
                                    'Modified by Lydia 2018/11/08 如果遇到跳頁,無法直接跳到下一列
                                    '.Selection.MoveDown Unit:=wdLine, Count:=1
                                    .Selection.SelectColumn
                                    .Selection.Collapse Direction:=wdCollapseEnd
                                    .Selection.MoveLeft Unit:=wdCell, Count:=1
                                End If
                             End If
                             mRg1 = intJ
                             strGrp = m_Item(iCall).IA01
                         End If
                         .Selection.MoveRight Unit:=wdCell, Count:=1
                         .Selection.Font.Shading.BackgroundPatternColorIndex = wdNoHighlight
                         For intR = 3 To 10
                              strExc(1) = ReadItemWordArray(m_Item, iCall, intR)
                              .Selection.TypeText Text:=strExc(1)
                              .Selection.MoveRight Unit:=wdCell, Count:=1
                         Next intR
                         intJ = intJ + 1
                    Next iCall '讀取案件統計陣列
                    mRg2 = intJ - 1
                    If mRg2 > mRg1 Then  '最後-合併
                        Set mRgrng = .Selection.Tables(1).Cell(mRg1, 1).Range
                        mRgrng.End = .Selection.Tables(1).Cell(mRg2, 1).Range.End
                        mRgrng.Select
                        .Selection.Cells.Merge
                        .Selection.MoveDown Unit:=wdLine, Count:=1
                    End If
                Else '沒有案件統計
                    intJ = intJ + 1
                    .Selection.Paragraphs.Alignment = wdAlignParagraphCenter
                    .Selection.TypeText Text:="案件往來"
                    .Selection.MoveRight Unit:=wdCell, Count:=2
                End If  '案件往來
                '---------
                intJ = intJ + 1
                .Selection.SelectRow
                .Selection.Cells.Merge
                .Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
                .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(3), RulerStyle:=wdAdjustProportional
                .Selection.Collapse Direction:=wdCollapseStart
                .Selection.Paragraphs.Alignment = wdAlignParagraphCenter
                .Selection.TypeText Text:="拜訪記錄"
                .Selection.MoveRight Unit:=wdCell, Count:=1
                .Selection.Paragraphs.Alignment = wdAlignParagraphLeft
                '讀取-往來記錄(含關係企業)
                tmpArr = Empty
                tmpArr = Split(strFAlist, "；")
                strGrp = ""
                For iCall = 0 To UBound(tmpArr)
                    If Trim(tmpArr(iCall)) <> "" Then   '關係企業
                        'Modified by Lydia 2024/02/27 用往來排序; order by cr01 >> order by cr02
                        'Modified by Lydia 2024/09/02 往來紀錄中若有「B21 互惠：評估結果」預設為自動顯示在附件word中的第一項，同一Y/X編號只會用一筆B21進行評估
                        'strSql = "select * from contactrecord where cr03=" & CNULL(Mid(tmpArr(iCall), 1, 9)) & " order by cr02"
                        strSql = "select * from contactrecord where cr03=" & CNULL(Mid(tmpArr(iCall), 1, 9)) & " order by decode(cr05,'B21',0,1), cr02"
                        intR = 1
                        Set rsRD = ClsLawReadRstMsg(intR, strSql)
                        If intR = 1 Then '有往來記錄
                            intCR = intCR + rsRD.RecordCount 'Added by Lydia 2025/08/22
                            rsRD.MoveFirst
                            Do While Not rsRD.EOF
                                strExc(1) = "● "
                                strExc(2) = "　 "
                                strExc(1) = strExc(1) & "往來日期：" & ChangeWStringToWDateString("" & rsRD.Fields("cr02"))
                                If "" & rsRD.Fields("cr04") <> "" Then
                                    strExc(1) = strExc(1) & vbCrLf & strExc(2) & "連絡人： " & Trim("" & rsRD.Fields("cr04"))
                                End If
                                If "" & rsRD.Fields("cr05") <> "" Then
                                    strExc(1) = strExc(1) & vbCrLf & strExc(2) & "往來類別： " & Trim("" & rsRD.Fields("cr05"))
                                End If
                                If "" & rsRD.Fields("cr06") <> "" Then
                                    strExc(1) = strExc(1) & vbCrLf & strExc(2) & "主旨： " & Trim("" & rsRD.Fields("cr06"))
                                End If
                                If "" & rsRD.Fields("cr07") <> "" Then
                                    strExc(1) = strExc(1) & vbCrLf & strExc(2) & "地點： " & Trim("" & rsRD.Fields("cr07"))
                                End If
                                If "" & rsRD.Fields("cr19") <> "" Then
                                    strExc(3) = ""
                                    tmpArr2 = Empty
                                    tmpArr2 = Split("" & rsRD.Fields("cr19"), ",")
                                    For intR = 0 To UBound(tmpArr2)
                                        If Trim(tmpArr2(intR)) <> "" Then
                                            strExc(4) = GetJobTitle(Trim(tmpArr2(intR)), 2, True) '員工名稱+職稱
                                            strExc(3) = strExc(3) & "、" & Trim(strExc(4))
                                        End If
                                    Next intR
                                    strExc(1) = strExc(1) & vbCrLf & strExc(2) & "接洽同仁： " & Mid(strExc(3), 2)
                                End If
                                If "" & rsRD.Fields("cr08") <> "" Then
                                    strExc(1) = strExc(1) & vbCrLf & strExc(2) & "內容： " & Trim("" & rsRD.Fields("cr08"))
                                End If
                                '附件
                                If "" & rsRD.Fields("cr09") <> "" Then
                                    strExc(3) = ""
                                    tmpArr2 = Empty
                                    tmpArr2 = Split("" & rsRD.Fields("cr09"), ",")
                                    For intR = 0 To UBound(tmpArr2)
                                        If Trim(tmpArr2(intR)) <> "" Then
                                             strExc(4) = Trim(tmpArr2(intR))
                                             strExc(4) = Mid(strExc(4), 1, InStrRev(strExc(4), " ("))
                                             strExc(3) = strExc(3) & IIf(strExc(3) <> "", vbCrLf & strExc(2), "") & strExc(4)
                                        End If
                                    Next intR
                                    strExc(1) = strExc(1) & vbCrLf & strExc(2) & "附件： " & strExc(3)
                                End If
                                '代理人資料
                                If strGrp <> "***" & Trim(Mid(tmpArr(iCall), 11)) & "***" Then
                                    strExc(1) = "***" & Trim(Mid(tmpArr(iCall), 11)) & "***" & vbCrLf & strExc(1)
                                    strGrp = "***" & Trim(Mid(tmpArr(iCall), 11)) & "***"
                                End If
                                .Selection.TypeText Text:=strExc(1)
                                .Selection.TypeParagraph
                                 rsRD.MoveNext
                            Loop
                        End If '有往來記錄
                    End If '關係企業
                Next iCall
                '---------
          End With
          '最後處理文字格式
          For iCall = 0 To UBound(tmpArr)
               g_WordAp.Selection.WholeStory
               '代理人／廠商名稱
               WordFindText "***" & strTemp(1) & " " & strTemp(0) & "(" & pNo & ")***", "1", strTemp(1) & " " & strTemp(0) & "(" & pNo & ")"
               '往來記錄(名稱)
               If iUpper > 0 Then 'Added by Lydia 2023/01/11 判斷有案件統計
                   WordFindText "***" & m_Item(1).IA01 & "***", "1", m_Item(1).IA01
               End If 'Added by Lydia 2023/01/11
               If Trim(tmpArr(iCall)) <> "" Then   '關係企業
                   If Mid(tmpArr(iCall), 1, 9) = pNo Then
                        WordFindText "***" & Trim(Mid(tmpArr(iCall), 11)) & "***", "1", Trim(Mid(tmpArr(iCall), 11))
                   Else
                        WordFindText "***" & Trim(Mid(tmpArr(iCall), 11)) & "***", "2", Trim(Mid(tmpArr(iCall), 11))
                   End If
               End If
          Next iCall
          
          g_WordAp.Selection.HomeKey Unit:=wdStory
          'Modifie dby Lydia 2019/04/09 改成模組
          'g_WordAp.Visible = True
          'g_WordAp.WindowState = wdWindowStateMaximize
          bVisible = True '預設顯示Word
          Pub_RePosWord g_WordAp, bVisible, m_WordLeft, m_WordTop
          'end  2019/04/09
          Set g_WordAp = Nothing
          MsgBox "Word資料已產生完成 !", vbInformation

      End If 'Added by Lydia 2019/04/09 改成模組
    End If
    InsertQueryLog (intCR) 'Added by Lydia 2025/08/22
    
ErrHandle:
    Set rsRD = Nothing
    Exit Sub
    
    If Err.Number <> 0 Then
       Select Case Err.Number
          Case 91, 462:
             Set g_WordAp = New Word.Application
             g_WordAp.Documents.add
             If bolRetry = False Then
                bolRetry = True
                Resume
             End If
          Case Else:
             MsgBox "錯誤 : " & Err.Description, vbCritical
       End Select
    End If
End Sub


'讀取Word格式資料
Private Function ReadItemWordArray(ByRef pArray() As INVITEM, ByVal pRow As Integer, ByVal pInx As Integer) As String
Dim rData As String

    ReadItemWordArray = ""
    If pRow = 0 And pInx = 0 Then Exit Function
    
    Select Case pInx
          Case 1: rData = pArray(pRow).IA01
          Case 2: rData = pArray(pRow).IA02
          Case 3: rData = pArray(pRow).IA03
          Case 4: rData = pArray(pRow).IA04
          Case 5: rData = pArray(pRow).IA05
          Case 6: rData = pArray(pRow).IA06
          Case 7: rData = pArray(pRow).IA07
          Case 8: rData = pArray(pRow).IA08
          Case 9: rData = pArray(pRow).IA09
          Case 10: rData = pArray(pRow).IA10
    End Select
    ReadItemWordArray = rData
    
End Function

'儲存Word格式資料
Private Sub SetItemWordArray(ByRef pArray() As INVITEM, ByVal pRow As Integer, ByVal pInx As Integer, ByVal pData As String)
    
    Select Case pInx
          Case 1: pArray(pRow).IA01 = pData
          Case 2: pArray(pRow).IA02 = pData
          Case 3: pArray(pRow).IA03 = pData
          Case 4: pArray(pRow).IA04 = Val(pArray(pRow).IA04) + Val(pData)
          Case 5: pArray(pRow).IA05 = Val(pArray(pRow).IA05) + Val(pData)
          Case 6: pArray(pRow).IA06 = Val(pArray(pRow).IA06) + Val(pData)
          Case 7: pArray(pRow).IA07 = Val(pArray(pRow).IA07) + Val(pData)
          Case 8: pArray(pRow).IA08 = Val(pArray(pRow).IA08) + Val(pData)
          Case 9: pArray(pRow).IA09 = Val(pArray(pRow).IA09) + Val(pData)
          Case 10: pArray(pRow).IA10 = Val(pArray(pRow).IA10) + Val(pData)
    End Select
End Sub

'尋找Word檔中文字
Private Sub WordFindText(ByVal strFindText As String, ByVal nType As String, Optional strReplaceText As String = "")
   If Trim(strFindText) = "" Then Exit Sub
   With g_WordAp
      .Selection.Find.ClearFormatting
      .Selection.Find.Text = strFindText
      .Selection.Find.Replacement.Text = ""
      .Selection.Find.Forward = True
      .Selection.Find.Wrap = wdFindContinue
      .Selection.Find.Format = False
      .Selection.Find.MatchCase = False
      .Selection.Find.MatchWholeWord = False
      .Selection.Find.MatchWildcards = False
      .Selection.Find.MatchSoundsLike = False
      .Selection.Find.MatchAllWordForms = False
      .Selection.Find.MatchByte = True
      If .Selection.Find.Execute = True Then
            .Selection.Delete
            .Selection.Font.ColorIndex = wdBlack
            If nType = "1" Then '顯示HightLine(黃)
                  .Selection.Font.Shading.BackgroundPatternColorIndex = wdYellow
            ElseIf nType = "2" Then  '顯示HightLine(灰)
                  .Selection.Font.Shading.BackgroundPatternColorIndex = wdGray25
            End If
            .Selection.TypeText strReplaceText
      End If
   End With
End Sub

'Added by Lydia 2018/11/07 排除特定案件性質(解除期限、取消收文和更換代理人)
'Mark by Lydia 2025/09/18 不用了
'Private Function GetCaseClosePtyList(ByVal iKind As Integer, Optional ByRef iSysNo As String = "", Optional ByVal bolMerge As Boolean = True) As String
''iKind : 系統種類(1-->專  2-->商  3-->法  4-->顧  5-->服)
''iSysNo: 系統別
''bolMerge : True=回傳系統別+案件性質
'Dim stMid01 As String, stMid02 As String, stMid03 As String
'Dim intM As Integer, tmpA1 As Variant
'
'If iKind = 0 Then Exit Function
'
'    If iSysNo = "" Then
'        iSysNo = SQLGrpStr("", iKind)
'    End If
'    If iSysNo <> "" Then
'        Select Case iKind
'            Case 1 '專利
'                  stMid01 = "907,913,925,937,902"
'            Case 2 '商標
'                  stMid01 = "703,704,718,726,720"
'            Case 3, 4 '法務,顧問
'                  stMid01 = "991,993,992,994"
'            Case 5 '服務
'                  stMid01 = ""
'        End Select
'        iSysNo = Replace(iSysNo, ",' '", "")  '清除空白系統別
'        If stMid01 <> "" Then
'            If bolMerge = True Then
'                If InStr(iSysNo, ",") = 0 Then
'                    stMid03 = GetAddStr(iSysNo) & ","
'                Else
'                    stMid03 = iSysNo & ","
'                End If
'                tmpA1 = Split(stMid01, ",")
'                For intM = 0 To UBound(tmpA1)
'                    If Trim(tmpA1(intM)) <> "" Then
'                        stMid02 = stMid02 & Replace(stMid03, "',", Trim(tmpA1(intM)) & "',")
'                    End If
'                Next intM
'                stMid02 = Mid(stMid02, 1, Len(stMid02) - 1)
'            Else
'                stMid02 = GetAddStr(stMid01)
'            End If
'        End If
'        GetCaseClosePtyList = stMid02
'    End If
'
'End Function
'end 2025/09/18

'Add by Amy 2020/10/15 勾選時判斷有往來記錄,往來記錄鈕變色
Private Sub ChkContactRecordBT(ByVal stChk As String, ByVal stKey As String)
    'Memo by Amy 2023/09/27  原2023/08/28 將按鈕鎖住,有資料才可按,User 按此鈕新增,故不鎖
    cmdOK(6).BackColor = &H8000000F
    If stChk = "V" And PUB_ChkContactRecord(stKey) = True Then
        cmdOK(6).BackColor = vbYellow
    End If
End Sub

'Added by Lydia 2021/01/05
Private Sub txtFM2_GotFocus(Index As Integer)
    TextInverse txtFM2(Index)
End Sub

Private Sub txtFM2_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtFM2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index = 0 Then
         Option1(1).Value = True
    End If
End Sub
'end 2022/01/5

'Add by Amy 2023/08/28 查詢只有一筆資料Grid顏色設定
Private Sub SetGridOneData()
   grdDataList.Visible = False
     With Me.grdDataList
        If .Rows = 2 Then
            .row = 1
            .col = 1
            If .Text <> "" Then
                .row = 1
                .col = 0
                .Text = "V"
                '變過色仍需要再跑,因為只有一筆選取時要變藍色
                Call SetMSGridColorQCus(1, Me.Name, grdDataList, strField, IIf(Check3.Value = vbChecked, True, False))
                '勾選時判斷有往來記錄,往來記錄鈕變色
                strExc(10) = grdDataList.TextMatrix(grdDataList.row, GetValue("編號"))
                If Left(strExc(10), 1) = "X" Or Left(strExc(10), 1) = "Y" Or Left(strExc(10), 1) = "R" Or Left(strExc(10), 2) = "平台" Then
                  Call ChkContactRecordBT(grdDataList.TextMatrix(grdDataList.row, GetValue("V")), strExc(10))
                End If
            End If
        End If
     End With
   grdDataList.Visible = True
End Sub


