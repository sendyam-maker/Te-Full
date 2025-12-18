VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100102_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "以申請人查詢"
   ClientHeight    =   5720
   ClientLeft      =   4010
   ClientTop       =   2450
   ClientWidth     =   9310
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5720
   ScaleWidth      =   9310
   Begin VB.CommandButton cmdOK 
      Caption         =   "卷宗區"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   10
      Left            =   4740
      Style           =   1  '圖片外觀
      TabIndex        =   18
      Top             =   30
      Width           =   720
   End
   Begin VB.CheckBox chkpct 
      Caption         =   "Check1"
      Height          =   255
      Left            =   7110
      TabIndex        =   15
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "代表圖"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   9
      Left            =   7110
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   30
      Width           =   720
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "顧問電話諮詢"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   8
      Left            =   30
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   30
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "專利相關案件"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   7
      Left            =   1200
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   30
      Width           =   1185
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "性質統計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   6
      Left            =   2400
      TabIndex        =   11
      Top             =   30
      Width           =   825
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   4455
      Left            =   60
      TabIndex        =   10
      Top             =   1230
      Width           =   9210
      _ExtentX        =   16228
      _ExtentY        =   7849
      _Version        =   393216
      Cols            =   13
      FixedCols       =   4
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
      _Band(0).Cols   =   13
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "子案"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   3
      Left            =   6450
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   30
      Width           =   645
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   5
      Left            =   8610
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   30
      Width           =   645
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "基本資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   0
      Left            =   3240
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   30
      Width           =   825
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   1
      Left            =   4080
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   30
      Width           =   645
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "   純關係   企業案"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   2
      Left            =   5490
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   30
      Width           =   945
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "下一筆"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   4
      Left            =   7860
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   30
      Width           =   720
   End
   Begin MSForms.Label Label7 
      Height          =   252
      Left            =   2280
      TabIndex        =   21
      Top             =   756
      Width           =   1188
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2090;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblContact 
      Height          =   255
      Left            =   7890
      TabIndex        =   20
      Top             =   480
      Visible         =   0   'False
      Width           =   1185
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2090;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label4 
      Height          =   252
      Left            =   2280
      TabIndex        =   19
      Top             =   480
      Width           =   4800
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "8467;444"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblContactL 
      AutoSize        =   -1  'True
      Caption         =   "接洽人："
      Height          =   255
      Left            =   7110
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "符號說明：＊閉卷;△非申請人案;●銷卷;◎商標舊委任狀;□商標新委任狀;＃台灣專利總委任書;e台灣電子證書"
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   60
      TabIndex        =   12
      Top             =   1050
      Width           =   8415
   End
   Begin VB.Label Label6 
      Caption         =   " "
      Height          =   255
      Left            =   1140
      TabIndex        =   3
      Top             =   750
      Width           =   1050
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   255
      Left            =   30
      TabIndex        =   2
      Top             =   750
      Width           =   900
   End
   Begin VB.Label Label3 
      Caption         =   " "
      Height          =   255
      Left            =   1140
      TabIndex        =   1
      Top             =   480
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人編號："
      Height          =   255
      Left            =   30
      TabIndex        =   0
      Top             =   480
      Width           =   1080
   End
End
Attribute VB_Name = "frm100102_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/16 改成Form2.0 ; GrdDataList改字型=新細明體-ExtB、Label4、Label7、lblContact
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/26 日期欄已修改
'Memo by Lydia 2020/09/17 cmdOK(2).caption從「關係企業案」改為「純關係企業案」與國外關聯企業做區隔；案件只抓母號前6碼相同的關係企業案
'2005/10/04  Nickc 重整修正
Option Explicit

Dim strSQL1 As String
Dim strSql  As String
Dim StrTag As String, i As Integer, j As Integer, s As Integer, intK As Integer, strTemp As Variant
Dim Str02 As String, Str03 As String, Str04 As String, Str05 As String, Str06 As String, Str07 As String
Dim Str01 As String
Dim strArr(62) As String, StrOk(32) As String, StrOkTxt(12) As String
Dim BolFrom100114 As Boolean
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
'Add by Morgan 2003/12/22 抓最後的繳費年度' ',substr(lpad(pa72,200,' '),200,1),
'edit by nickc 2007/01/23  少判斷第一年的  ' ',substr(lpad(pa72,200,' '),200,1),
'Private Const cntLstPayYearSQL As String = " decode( substr(lpad(pa72,200,' '),200,1),' ',' ',decode( substr(lpad(pa72,200,' '),199,1),',',substr(lpad(pa72,200,' '),200,1) ,decode( substr(lpad(pa72,200,' '),198,1),',',substr(lpad(pa72,200,' '),199,2) ,decode( substr(lpad(pa72,200,' '),197,1),',',substr(lpad(pa72,200,' '),198,3) ,decode( substr(lpad(pa72,200,' '),196,1),',',substr(lpad(pa72,200,' '),197,4) ) ) ) ) )"
Private Const cntLstPayYearSQL As String = " decode( substr(lpad(pa72,200,' '),200,1),' ',' ',decode( substr(lpad(pa72,200,' '),199,1),' ',substr(lpad(pa72,200,' '),200,1),',',substr(lpad(pa72,200,' '),200,1) ,decode( substr(lpad(pa72,200,' '),198,1),',',substr(lpad(pa72,200,' '),199,2) ,decode( substr(lpad(pa72,200,' '),197,1),',',substr(lpad(pa72,200,' '),198,3) ,decode( substr(lpad(pa72,200,' '),196,1),',',substr(lpad(pa72,200,' '),197,4) ) ) ) ) )"
'add by  nickc 2005/10/04  判斷是否法務
Public bolIsL As Boolean
'add by nickc 2005/10/04 若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
Private Const cntFaSql As String = " DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65))"
'Add by Morgan 2008/11/27
'為了要能共用，前畫面條件改以參數方式傳遞
Public m_Sys As String '系統類別
Public m_Cty1 As String, m_Cty2 As String '申請國家
Public m_Pty1 As String, m_Pty2 As String '案件性質
Public m_Type As String '收發文別
Public m_Date1 As String, m_Date2 As String '日期
Public m_CKind As String '是否含C類來函 N:不含
Public m_CaseNo As String '本所案號 for strMenu4
'Added by Lydia 2018/02/09
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim m_adoRst As ADODB.Recordset
'Added by Lydia 2019/11/01 利益衝突案件-管制
'Mark by Lydia 2019/12/26
'Dim m_CuFaArea As String '利益衝突案件：XY特殊權限管制-系統別
'Dim m_CuFaRight As String '利益衝突案件：XY特殊權限管制-可使用系統別
'Dim stCuFaSQL As String '利益衝突案件：查詢權限內的案件SQL
'Dim stConPA As String   '組合條件(Patent)
'Dim stConSP As String   '組合條件(ServicePractice)
'Dim rsCnt As New ADODB.Recordset
'end 2019/12/26
'Memo by Lydia 2019/12/26 利益衝突案件：於後面增加欄位; 從外層SQL控制，改成逐案比對。
Dim intCufaCnt As Integer '限閱案件X件
Dim m_AllSys As String
Dim SeColPA As String
Dim SeColTM As String
Dim SeColSP As String
Dim SeColLC As String
Dim SeColHC As String
Dim strSQLE(1) As String, strSqlEW(1) As String 'Add by Amy 2022/12/14
Dim strEField(1) As String 'Add by Amy 2023/03/06
Dim m_pub_QL05 As String 'Add By Sindy 2025/8/28 只記錄於此Form


Private Sub SetDataListWidth()
'Add by Morgan 2003/12/18
Dim arrGridHeadText, arrGridHeadWidth, iDep As String
Dim iCol As Integer
'edit by nickc 2005/05/06
'arrGridHeadText = Array("V", "本所案號", "分所號", "案件名稱", "申請國家" _
                     , "申請案號", "申請日", "審定專利號數", "准駁", "申請人1" _
                     , "商品類別", "專用期間", "專利公告號", "最後已繳年度", "申請人2" _
                     , "申請人3", "申請人4", "申請人5", "")
iDep = PUB_GetST06(strUserNum)

If bolIsL = False Then
   'edit by nickc 2007/03/23
   'edit by nickc 2007/12/21
   'If frm100102_1.chkpct.Value = vbChecked Then
   If ChkPCT.Value = vbChecked Then
       'Modified by Lydia 2019/12/26 +申請人1~5(cust01~cust05),FC代理人;
       'arrGridHeadText = Array("V", "本所案號", "分所號", "案件名稱", "申請國家" _
                      , "申請案號", "申請日", "PCT", "准駁", "申請人1" _
                      , "商品類別", "專用期間", "專利公告號", "最後已繳年度", "申請人2" _
                      , "申請人3", "申請人4", "申請人5", "", "")
       arrGridHeadText = Array("V", "本所案號", "分所號", "案件名稱", "申請國家" _
                      , "申請案號", "申請日", "PCT", "准駁", "申請人1" _
                      , "商品類別", "專用期間", "專利公告號", "最後已繳年度", "申請人2" _
                      , "申請人3", "申請人4", "申請人5", "FSORT", "CNT" _
                      , "CUST01", "CUST02", "CUST03", "CUST04", "CUST05", "FCNO")
   Else
       'Modified by Lydia 2019/12/26 +申請人1~5(cust01~cust05),FC代理人;
       'arrGridHeadText = Array("V", "本所案號", "分所號", "案件名稱", "申請國家" _
                      , "申請案號", "申請日", "審定專利號數", "准駁", "申請人1" _
                      , "商品類別", "專用期間", "專利公告號", "最後已繳年度", "申請人2" _
                      , "申請人3", "申請人4", "申請人5", "", "")
       arrGridHeadText = Array("V", "本所案號", "分所號", "案件名稱", "申請國家" _
                      , "申請案號", "申請日", "審定專利號數", "准駁", "申請人1" _
                      , "商品類別", "專用期間", "專利公告號", "最後已繳年度", "申請人2" _
                      , "申請人3", "申請人4", "申請人5", "FSORT", "CNT" _
                      , "CUST01", "CUST02", "CUST03", "CUST04", "CUST05", "FCNO")
   End If
'edit by nickc 2005/05/06
'   arrGridHeadWidth = Array(200, 1600, 0, 1600, 800 _
                     , 1200, 720, 1150, 450, 1000 _
                     , 800, 1800, 1000, 1200, 1100 _
                     , 1100, 1100, 1100, 1000)
   'Modify by Morgan 2004/2/27
   '電腦中心，跟分所才秀
   If PUB_GetST03(strUserNum) <> "M51" And iDep = "1" Then
        'edit by nickc 2007/03/23
        'edit by nickc 2007/12/21
        'If frm100102_1.chkpct.Value = vbChecked Then
        If ChkPCT.Value = vbChecked Then
            'Modified by Lydia 2019/12/26 +申請人1~5(cust01~cust05),FC代理人;
            arrGridHeadWidth = Array(200, 1600, 0, 1600, 800 _
                              , 1200, 800, 620, 450, 1000 _
                              , 800, 1800, 1000, 1200, 1100 _
                              , 1100, 1100, 1100, 0, 0 _
                              , 0, 0, 0, 0, 0, 0)
       Else
            'Modified by Lydia 2019/12/26 +申請人1~5(cust01~cust05),FC代理人;
            arrGridHeadWidth = Array(200, 1600, 0, 1600, 800 _
                              , 1200, 800, 1150, 450, 1000 _
                              , 800, 1800, 1000, 1200, 1100 _
                              , 1100, 1100, 1100, 0, 0 _
                              , 0, 0, 0, 0, 0, 0)
       End If
   Else
        'edit by nickc 2007/03/23
        'edit by nickc 2007/12/21
        'If frm100102_1.chkpct.Value = vbChecked Then
        If ChkPCT.Value = vbChecked Then
                'Modified by Lydia 2019/12/26 +申請人1~5(cust01~cust05),FC代理人;
                arrGridHeadWidth = Array(200, 1600, 900, 1600, 800 _
                                  , 1200, 800, 440, 450, 1000 _
                                  , 800, 1800, 1000, 1200, 1100 _
                                  , 1100, 1100, 1100, 0, 0 _
                                  , 0, 0, 0, 0, 0, 0)
        Else
                'Modified by Lydia 2019/12/26 +申請人1~5(cust01~cust05),FC代理人;
                arrGridHeadWidth = Array(200, 1600, 900, 1600, 800 _
                                  , 1200, 800, 1200, 450, 1000 _
                                  , 800, 1800, 1000, 1200, 1100 _
                                  , 1100, 1100, 1100, 0, 0 _
                                  , 0, 0, 0, 0, 0, 0)
        End If
   End If
Else
   '2005/11/29 MODIFY BY SONIA 申請國家縮小,顧問期間改放在進度備註欄,改案件性質,相關人欄位位置
   '2005/12/19 MODIFY BY SONIA 進度備註改抓案件名稱,取消結果欄,增回執日欄,相關人改相對人,依立卷問題3需求調整欄位位置
   'Modified by Lydia 2015/10/05 '承辦律師'改為'承辦人'、'承辦法務'改為'協辦人員'
   'Modified by Lydia 2019/12/26 +申請人1~5(cust01~cust05),FC代理人;
   'arrGridHeadText = Array("V", "本所案號", "分所號", "案件名稱", "相對人" _
                              , "收文日", "承辦人", "協辦人員", "發文日" _
                              , "回執日", "智權人員", "案件性質", "國家" _
                              , "取消收文日", "代理人", "", "")
   arrGridHeadText = Array("V", "本所案號", "分所號", "案件名稱", "相對人" _
                              , "收文日", "承辦人", "協辦人員", "發文日" _
                              , "回執日", "智權人員", "案件性質", "國家" _
                              , "取消收文日", "代理人", "FSORT", "CP09", "CNT" _
                              , "CUST01", "CUST02", "CUST03", "CUST04", "CUST05", "FCNO")
   '電腦中心，跟分所才秀
   If PUB_GetST03(strUserNum) <> "M51" And iDep = "1" Then
         'Modified by Lydia 2019/12/26 +申請人1~5(cust01~cust05),FC代理人;
         'arrGridHeadWidth = Array(200, 1400, 0, 1500, 600 _
                                , 800, 800, 800, 800 _
                                , 800, 600, 800, 500 _
                                , 1000, 1200, 0, 0 _
                                , 0, 0, 0, 0, 0, 0)
         arrGridHeadWidth = Array(200, 1400, 0, 1500, 600 _
                                , 800, 800, 800, 800 _
                                , 800, 600, 800, 500 _
                                , 1000, 1200, 0, 0, 0 _
                                , 0, 0, 0, 0, 0, 0)
   Else
         'Modified by Lydia 2019/12/26 +申請人1~5(cust01~cust05),FC代理人;
         'arrGridHeadWidth = Array(200, 1400, 900, 1500, 600 _
                                , 800, 800, 800, 800 _
                                , 800, 600, 800, 500 _
                                , 1000, 1200, 0, 0 _
                                , 0, 0, 0, 0, 0, 0)
         arrGridHeadWidth = Array(200, 1400, 900, 1500, 600 _
                                , 800, 800, 800, 800 _
                                , 800, 600, 800, 500 _
                                , 1000, 1200, 0, 0, 0 _
                                , 0, 0, 0, 0, 0, 0)
   End If
End If

grdDataList.Cols = UBound(arrGridHeadText) + 1
For iCol = 0 To grdDataList.Cols - 1
   'add by nick 2004/07/07
   grdDataList.row = 0
   grdDataList.col = iCol
   grdDataList.Text = arrGridHeadText(iCol)
   grdDataList.ColWidth(iCol) = arrGridHeadWidth(iCol)
   grdDataList.CellAlignment = flexAlignCenterCenter
Next iCol
End Sub

Private Sub cmdcp10_Click()
'92.04.16 nick 紀錄作用按鍵
cmdState = 6
PubShowNextData
Exit Sub
End Sub

'92.04.16 nick
Public Sub PubShowNextData()
'2cmd
Select Case cmdState
Case 0 '案件基本資料
      Me.Enabled = False
      For i = 1 To grdDataList.Rows - 1
         grdDataList.col = 0
         grdDataList.row = i
         If Trim(grdDataList.Text) = "V" Then
            grdDataList.col = 0
            grdDataList.Text = ""
            'Add by Morgan 2004/4/9
            '固定欄位不變色
            'For j = 0 To grdDataList.Cols - 1
            For j = 4 To grdDataList.Cols - 1
                 grdDataList.col = j
                 grdDataList.CellBackColor = QBColor(15)
            Next j
           Dim Str01 As String
           grdDataList.col = 1
           Str01 = SystemNumber(Replace(grdDataList, "△", ""), 1)
           If Mid(UCase(Str01), 1, 1) = "N" Then
               Str01 = Mid(Str01, 2, 3)
           End If
           If Not IsNull(grdDataList.Text) Then
               If fnSaveParentForm(Me) = False Then
                   Me.Enabled = True
                   Exit Sub
               End If
                Select Case Pub_RplStr(Str01)
                    Case "CFP", "FCP", "P"   '專利
                          Screen.MousePointer = vbHourglass
                          frm100101_3.Show
                          frm100101_3.Tag = Pub_RplStr(grdDataList.Text)
                          frm100101_3.StrMenu
                          Screen.MousePointer = vbDefault
                    Case "CFT", "FCT", "T", "TF"   '商標
                          Screen.MousePointer = vbHourglass
                          frm100101_4.Show
                          frm100101_4.Tag = Pub_RplStr(grdDataList.Text)
                          frm100101_4.StrMenu
                          Screen.MousePointer = vbDefault
                    'Modify By Sindy 2009/07/24 增加LIN系統類別
                    'modify by sonia 2019/7/29 +ACS系統類別
                    Case "CFL", "FCL", "L", "LIN", "ACS"     '法務
                          Screen.MousePointer = vbHourglass
                          frm100101_5.Show
                          frm100101_5.Tag = Pub_RplStr(grdDataList.Text)
                          frm100101_5.StrMenu
                          Screen.MousePointer = vbDefault
                    Case "LA"            '顧問
                          Screen.MousePointer = vbHourglass
                          frm100101_6.Show
                          frm100101_6.Tag = Pub_RplStr(grdDataList.Text)
                          frm100101_6.StrMenu
                          Screen.MousePointer = vbDefault
                    Case Else                  '服務
                         Select Case Pub_RplStr(Str01)
                             Case "TB"    '條碼
                                 Screen.MousePointer = vbHourglass
                                 frm100101_7.Show
                                 frm100101_7.Tag = Pub_RplStr(grdDataList.Text)
                                 frm100101_7.StrMenu
                                 Screen.MousePointer = vbDefault
                             Case "TM"
                                 Screen.MousePointer = vbHourglass
                                 frm100101_8.Show
                                 frm100101_8.Tag = Pub_RplStr(grdDataList.Text)
                                 frm100101_8.StrMenu
                                 Screen.MousePointer = vbDefault
                             Case "TD"
                                 Screen.MousePointer = vbHourglass
                                 frm100101_9.Show
                                 frm100101_9.Tag = Pub_RplStr(grdDataList.Text)
                                 frm100101_9.StrMenu
                                 Screen.MousePointer = vbDefault
                             Case "TC", "CFC"
                                 Screen.MousePointer = vbHourglass
                                 frm100101_A.Show
                                 frm100101_A.Tag = Pub_RplStr(grdDataList.Text)
                                 frm100101_A.StrMenu
                                 Screen.MousePointer = vbDefault
                             Case Else
                                 Screen.MousePointer = vbHourglass
                                 frm100101_B.Show
                                 frm100101_B.Tag = Pub_RplStr(grdDataList.Text)
                                 frm100101_B.StrMenu
                                 Screen.MousePointer = vbDefault
                          End Select
                End Select
                 Me.Enabled = True
                 Exit Sub
           End If
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
            'Add by Morgan 2004/4/9
            '固定欄位不變色
            'For j = 0 To grdDataList.Cols - 1
            For j = 4 To grdDataList.Cols - 1
               grdDataList.col = j
               grdDataList.CellBackColor = QBColor(15)
            Next j
            grdDataList.col = 1
            If Not IsNull(grdDataList.Text) Then
               Screen.MousePointer = vbHourglass
                If fnSaveParentForm(Me) = False Then
                    Me.Enabled = True
                    Exit Sub
                End If
               frm100101_2.Show
               frm100101_2.Tag = Pub_RplStr(grdDataList.Text)
               frm100101_2.Label15.Caption = grdDataList.TextMatrix(i, 2)
               If BolFrom100114 = False Then
                  'Modify By Sindy 2021/4/21 不要為了排除C類而寫2個函數 StrMenu,StrMenu1(Mark)
'                  If Len(Trim(m_CKind)) = 0 Then
'                      frm100101_2.StrMenu
'                  Else
'                      frm100101_2.StrMenu1
'                  End If
                  frm100101_2.m_CKind = Me.m_CKind
                  frm100101_2.StrMenu
                  '2021/4/21 END
                  Screen.MousePointer = vbDefault
                  Me.Enabled = True
                  Exit Sub
               Else
                   frm100101_2.StrMenu
               End If
               frm100101_2.cmdok(0).Enabled = False
               frm100101_2.cmdok(1).Enabled = False
               Screen.MousePointer = vbDefault
               Me.Enabled = True
               Exit Sub
            End If
        End If
     Next i
     Me.Enabled = True
Case 2 '關係企業案件
     Call StrMenu1
Case 3 '子案資料
     Me.Enabled = False
     StrTag = ""
     For i = 1 To grdDataList.Rows - 1
     grdDataList.col = 0
     grdDataList.row = i
     If Trim(grdDataList.Text) = "V" Then
        grdDataList.col = 0
        grdDataList.Text = ""
         'Add by Morgan 2004/4/9
         '固定欄位不變色
         'For j = 0 To grdDataList.Cols - 1
        For j = 4 To grdDataList.Cols - 1
           grdDataList.col = j
           grdDataList.CellBackColor = QBColor(15)
        Next j
         grdDataList.col = 1
         If Not IsNull(grdDataList.Text) Then
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            frm100102_3.Show
            frm100102_3.Tag = Pub_RplStr(grdDataList.Text)
            frm100102_3.Label3.Caption = Me.Tag
            frm100102_3.Label7.Caption = Me.Label4.Caption
            frm100102_3.StrMenu
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Exit Sub
         End If
     End If
     Next i
     Me.Enabled = True
Case 4 '下一筆
     tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 5 '結束
     fnCloseAllFrm100
Case 6
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
    Screen.MousePointer = vbHourglass
    frm100114_3.Show
    frm100114_3.StrMenu Label3.Caption, BolFrom100114
    Screen.MousePointer = vbDefault
    Me.Enabled = True
Case 7
     Me.Enabled = False
     StrTag = ""
     For i = 1 To grdDataList.Rows - 1
     grdDataList.col = 0
     grdDataList.row = i
     If Trim(grdDataList.Text) = "V" Then
        grdDataList.col = 0
        grdDataList.Text = ""
         'Add by Morgan 2004/4/9
         '固定欄位不變色
         'For j = 0 To grdDataList.Cols - 1
        For j = 4 To grdDataList.Cols - 1
           grdDataList.col = j
           grdDataList.CellBackColor = QBColor(15)
        Next j
         grdDataList.col = 1
         If Not IsNull(grdDataList.Text) Then
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            frm100101_h.Show
            frm100101_h.KeyString = Pub_RplStr(grdDataList.Text)
            frm100101_h.SearchKind = "本所案號"
            frm100101_h.StrMenu
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Exit Sub
         End If
     End If
     Next i
     Me.Enabled = True
'2006/1/2 ADD BY SONIA
Case 8
   Me.Enabled = False
   For i = 1 To grdDataList.Rows - 1
      grdDataList.col = 0
      grdDataList.row = i
      If Trim(grdDataList.Text) = "V" Then
         grdDataList.col = 1
         StrTag = Pub_RplStr(grdDataList.Text)
         If SystemNumber(StrTag, 1) <> "LA" Then
            MsgBox ("點選資料非顧問案件, 不可按電話諮詢按鈕！")
            Me.Enabled = True
            Exit Sub
         End If
         grdDataList.col = 0
         grdDataList.Text = ""
         For j = 4 To grdDataList.Cols - 1
            grdDataList.col = j
            grdDataList.CellBackColor = QBColor(15)
         Next j
         grdDataList.col = 16
         If Not IsNull(grdDataList.Text) Then
            Set frm010015.UpForm = Me
            '2011/5/26 ADD BY SONIA
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            '2011/5/26 END
            Screen.MousePointer = vbHourglass
            frm010015.Show
            frm010015.Tag = grdDataList.Text & ChangeWStringToTString(GetTodayDate)
            frm010015.GetData (0)
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Exit Sub
         End If
      End If
   Next i
   Me.Enabled = True
'2006/1/2 END
'add by nickc 2007/09/05 薛說要加入代表圖按鈕
Case 9
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
            grdDataList.CellBackColor = QBColor(15)
        Next j
         grdDataList.col = 1
         If Not IsNull(grdDataList.Text) Then
            Me.Hide
            Screen.MousePointer = vbHourglass
            frmPic001.oCP01 = SystemNumber(Pub_RplStr(grdDataList.Text), 1)
            frmPic001.oCP02 = SystemNumber(Pub_RplStr(grdDataList.Text), 2)
            frmPic001.oCP03 = SystemNumber(Pub_RplStr(grdDataList.Text), 3)
            frmPic001.oCP04 = SystemNumber(Pub_RplStr(grdDataList.Text), 4)
            frmPic001.StrMenu
            'Mark by Amy 2018/07/18 開放可維護-秀玲
'            frmPic001.cmdok(0).Visible = False
'            frmPic001.cmdok(1).Visible = False
'            frmPic001.cmdok(2).Visible = False
'            frmPic001.cmdok(4).Visible = False
'            frmPic001.cmdok(5).Visible = False
'            frmPic001.cmdok(6).Visible = False
'            frmPic001.Label12.Visible = False
            'end 2018/07/18
            frmPic001.SetSeekCmdok 'Add by Amy 2018/07/18
            frmPic001.Show vbModal
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Me.Show
         End If
     End If
     Next i
     Me.Enabled = True
'Add By Sindy 2019/1/15
Case 10 '卷宗區
   Me.Enabled = False
   For i = 1 To grdDataList.Rows - 1
      grdDataList.col = 0
      grdDataList.row = i
      If Trim(grdDataList.Text) = "V" Then
         grdDataList.col = 1
         StrTag = Pub_RplStr(grdDataList.Text)
         grdDataList.col = 0
         grdDataList.Text = ""
         For j = 4 To grdDataList.Cols - 1
            grdDataList.col = j
            grdDataList.CellBackColor = QBColor(15)
         Next j
         grdDataList.col = 16
         If Not IsNull(grdDataList.Text) Then
            Set frm010015.UpForm = Me
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            frm100101_L.m_strKey = StrTag
            'frm100101_L.Hide
            frm100101_L.SetParent Me
            If frm100101_L.QueryData = True Then
               frm100101_L.Show
               Me.Hide
            End If
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Exit Sub
         End If
      End If
   Next i
   Me.Enabled = True
Case Else
End Select
End Sub

Public Sub cmdok_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
Exit Sub
End Sub

Private Sub Form_Activate()
   pub_QL05 = m_pub_QL05 'Add By Sindy 2025/8/28 還原此Form的查詢條件記錄
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   SetDataListWidth
   cmdState = -1
   bolIsL = False
   m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28 記錄此Form的查詢條件
End Sub

'Mark by Amy 2023/01/09 語法改共用函數
Sub StrMenu_Old()
'Dim strAppNo As String, strAppNo1 As String
'Dim strContactNo As String
'Dim strWhereCP As String 'Add by Amy 2022/11/14
'
'BolFrom100114 = False
'Me.Enabled = False
'Str01 = ""    '申請人編號
'Str02 = ""    '系統類別
'Str03 = ""    '收文日期(起)
'Str04 = ""    '收文日期(迄)
'Str05 = ""    '案件性質(起)
'Str06 = ""    '案件性質(迄)
'Str07 = ""    '是否含來函資料
'Str01 = Me.Tag
'
''Add By Sindy 2011/01/03 檢查國內外權限
'If CheckSR12(Str01) = False Then
'   Screen.MousePointer = vbDefault
'   Me.Enabled = True
'   tmpBol = fnCancelNowFormAndShowParentForm(Me)
'   Exit Sub
'End If
'
''2005/12/19 ADD BY SONIA
'If bolIsL = False Then
'   Label1 = "申請人編號："
'Else
'   Label1 = "當事人編號："
'End If
'
'lblContact.Caption = "" 'Added by Lydia 2021/12/16
'
''2005/12/19 END
'Str02 = IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys))
''收文日起
'Str03 = m_Date1
''收文日迄
'Str04 = IIf(Len(m_Date1) <= 0, "", IIf(Len(m_Date2) <= 0, (ServerDate - 19110000), m_Date2))
'Str05 = m_Pty1
'Str06 = m_Pty2
'Str07 = m_CKind
''組字串
'strSQL1 = ""
''收文
'If m_Type = "1" Then
'   If Len(Str03) <> 0 Then
'       strSQL1 = strSQL1 + " and cp05>=" & Val(ChangeTStringToWString(Str03))
'   End If
'   If Len(Str04) <> 0 Then
'       strSQL1 = strSQL1 + " and cp05<=" & Val(ChangeTStringToWString(Str04))
'   End If
''Add by Morgan 2008/11/26 配合代理人查詢畫面的條件
''發文
'Else
'   If Len(Str03) <> 0 Then
'       strSQL1 = strSQL1 + " and cp27>=" & Val(ChangeTStringToWString(Str03))
'   End If
'   If Len(Str04) <> 0 Then
'       strSQL1 = strSQL1 + " and cp27<=" & Val(ChangeTStringToWString(Str04))
'   End If
'End If
'
'If Len(Str05) <> 0 Then
'    strSQL1 = strSQL1 + " and cp10>='" & Str05 & "' "
'End If
'If Len(Str06) <> 0 Then
'    strSQL1 = strSQL1 + " and cp10<='" & Str06 & "' "
'End If
'If UCase(Str07) = "N" Then
'    strSQL1 = strSQL1 + " and cp09 < 'C' "
'End If
'
''Added by Lydia 2019/11/01 非法務案+屬於利益衝突案件之XY編號
''Mark by Lydia 2019/12/26
''stConPA = "": stConSP = ""
''If bolIsL = False And strSrvDate(1) >= XY特殊權限啟用日 And InStr(XY特殊權限範圍, Left(ChangeCustomerL(Me.Tag), 8)) > 0 Then
''    cnnConnection.Execute "delete from R100102_2 where R02201='" & strUserNum & "' and R02202='" & Me.Name & "' " '清空暫存檔
''    If PUB_ChkCuFa_Right(Me.Name, Me.Tag, Str02, m_CuFaRight, m_CuFaArea) = True Then
''    End If
''    '有管制系統別=>組合SQL條件
''    If m_CuFaArea <> "" Then
''        stConPA = Pub_CufaConSQL(Me.Name, "PA", Me.Tag, m_CuFaRight, m_CuFaArea)
''        stConSP = Pub_CufaConSQL(Me.Name, "SP", Me.Tag, m_CuFaRight, m_CuFaArea)
''    End If
''End If
''end 2019/11/01
''end 2019/12/26
'
''顯示表單上面的值
'Label3.Caption = Me.Tag
'
''Modify by Morgan 2008/8/12 考慮含接洽人編號
'If Mid(Me.Tag, 10, 1) = "-" Then
'   strContactNo = Mid(Me.Tag, 11)
'   Me.Tag = Left(Me.Tag, 9)
'   strAppNo = Me.Tag
'   strAppNo1 = Left(strAppNo, 8) & "9"
'Else
'   Me.Tag = ChangeCustomerL(Me.Tag)
'   strAppNo = Me.Tag
'   strAppNo1 = strAppNo
'End If
''Add by Amy 2022/11/14 X編號抓進度檔的CP55/CP56/CP89~CP96的資料時,要加入條件(CP158>0 OR CP159=0),否則未發文CP158=0且已取消收文CP159>0也會出現 ex:X28819010查案件,不應該帶出CFT-011352-秀玲
''Modify by Amy X75109010不應查出P-115659/P-115660,已發文但是後來有取消收文
''strWhereCP = " And (CP158>0 OR CP159=0) "
'strWhereCP = " And CP159=0 "
'
'If strContactNo <> "" Then
'   strSql = "SELECT NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),CU13,ST02,cu111,PCC05" & _
'      " FROM CUSTOMER,STAFF,POTCUSTCONT WHERE CU01='" & Left$(Me.Tag, 8) & "' AND CU02='" & Right$(Me.Tag, 1) & "' AND CU13=ST01(+)" & _
'      " AND PCC01(+)=CU01 AND PCC02(+)='" & strContactNo & "' "
'Else
'   strSql = "SELECT NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),CU13,ST02,cu111,'' PCC05" & _
'      " FROM CUSTOMER,STAFF WHERE CU01='" & Left$(Me.Tag, 8) & "' AND CU02='" & Right(Me.Tag, 1) & "' AND CU13=ST01(+) "
'End If
''end 2008/8/12
'
'CheckOC
'adoRecordset.CursorLocation = adUseClient
'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'    If IsNull(adoRecordset.Fields(0)) Then
'        Label4.Caption = ""
'    Else
'        Label4.Caption = adoRecordset.Fields(0)
'    End If
'    If IsNull(adoRecordset.Fields(1)) Then
'        Label6.Caption = ""
'    Else
'        Label6.Caption = adoRecordset.Fields(1)
'    End If
'    If IsNull(adoRecordset.Fields(2)) Then
'        Label7.Caption = ""
'    Else
'        Label7.Caption = adoRecordset.Fields(2)
'    End If
'    'add by nickc 2005/12/06
'    If CheckStr(adoRecordset.Fields("cu111")) = "Y" Then
'        Label3.ForeColor = &HFF&
'    Else
'        Label3.ForeColor = &H80000012
'    End If
'   'Add by Morgan 2008/8/12
'   If strContactNo <> "" Then
'      lblContactL.Visible = True
'      lblContact.Visible = True
'      lblContact = "" & adoRecordset.Fields("pcc05")
'   End If
'   'end 2008/8/12
'End If
'CheckOC
'
'    'Added by Lydia 2019/12/26 利益衝突案件：於後面增加欄位
'    SeColTM = " ,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno "
'    SeColPA = " ,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno "
'    SeColSP = " ,sp08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno "
'    SeColLC = " ,lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno "
'    SeColHC = " ,hc05 as cust01,hc24 as cust02,hc25 as cust03,hc26 as cust04,hc27 as cust05,'' as fcno "
'    'end 2019/12/26
'    'Add by Amy 2022/12/14 +e符號
'    'T、FCT已有專用期間者,或無專用期但進度檔商標之註冊費717已發文
'    strSQLE(0) = "Select tm01 as Ecp01,tm02 as Ecp02,tm03 as Ecp03,tm04 as Ecp04,'e' as EState From TradeMark " & _
'                        "Where tm136='1' And Tm04='00' And (tm21 is not null or (tm21 is null And Exists (Select * From CaseProgress Where cp01 in(" & SQLGrpStr2("", 2) & ") And cp10='717' And cp158<>0 )  ))"
'    strSqlEW(0) = " And tm01=Ecp01(+) And tm02=Ecp02(+) And tm03=Ecp03(+) And tm04=Ecp04(+) "
'    'P、FCP已有專用期間者,或無專用期但進度檔專利之領證601已發文
'    strSQLE(1) = "Select pa01 as Ecp01,pa02 as Ecp02,pa03 as Ecp03,pa04 as Ecp04,'e' as EState From Patent " & _
'                        "Where pa178='1' And pa04='00' And (pa24 is not null or (pa24 is null And Exists (Select * From CaseProgress Where cp01 in(" & SQLGrpStr2("", 1) & ") And cp10='601' And cp158<>0 )  ))"
'    strSqlEW(1) = " And pa01=Ecp01(+) And pa02=Ecp02(+) And pa03=Ecp03(+) And pa04=Ecp04(+) "
'    'end 2022/12/14
'
''add by nickc 2005/10/04
'If bolIsL = False Then '非法務案
''Add by Morgan 2003/12/18
''隱藏申請人1
'   GrdDataList.ColWidth(9) = 0
'
'      'add by nickc 2006/08/28  加入銷卷
'      'edit by nickc 2006/12/11 商標加申請人
'      'Modify by Morgan 2008/8/13 加查詢接洽人案件
'      'Modify By Sindy 2012/2/8 +||DECODE(TM128,'Y','◎','')
'      'Modify By Sindy 2012/5/29 DECODE(TM128,'Y','◎','')==>DECODE(TM128,'Y','◎','A','□','')
'      'Modified by Morgan 2019/1/30 SQLGrpStr(Str02, #)->GetAddStr(str02) 取消不必要系統別的檢查(減少語法執行次數,分所有較大影響)
'      'Modified by Lydia 2019/12/26 增加欄位SeColTM
'      'Modify by Amy 2022/12/14 +e符號 Nvl(EState,'')
'      strSql = "SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','')||Nvl(EState,'') AS 本所案號 ,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort,DECODE(TM123,NULL,C1.CU01||C1.CU127,C1.CU01||TM123) CNT" & SeColTM & _
'               " FROM TRADEMARK,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress,(" & strSQLE(0) & ") " & _
'               "WHERE tm10=na01(+) and " & IIf(strAppNo = strAppNo1, "TM23='" & strAppNo, "TM23>='" & strAppNo & "' and TM23<='" & strAppNo1) & "' and tm01 in (" & GetAddStr(Str02) & ") AND TM04='00' AND SUBSTR(TM23,1,8) = c1.CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1)) = c1.CU02(+) " & _
'               " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
'               " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) " & _
'               " and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) " & _
'               " and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) " & strSqlEW(0) & strSQL1
'      strSql = strSql + " union SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','')||Nvl(EState,'') AS 本所案號 ,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort,DECODE(TM123,NULL,C2.CU01||C2.CU127,C1.CU01||TM123) CNT" & SeColTM & _
'               " FROM TRADEMARK,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress,(" & strSQLE(0) & ") " & _
'               "WHERE tm10=na01(+) and " & IIf(strAppNo = strAppNo1, "TM78='" & strAppNo, "TM78>='" & strAppNo & "' and TM78<='" & strAppNo1) & "' and tm01 in (" & GetAddStr(Str02) & ") AND TM04='00' AND SUBSTR(TM23,1,8) = c1.CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1)) = c1.CU02(+) " & _
'               " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
'               " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) " & _
'               " and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) " & _
'               " and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) " & strSqlEW(0) & strSQL1
'      strSql = strSql + " union SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','')||Nvl(EState,'') AS 本所案號 ,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort,DECODE(TM123,NULL,C3.CU01||C3.CU127,C1.CU01||TM123) CNT" & SeColTM & _
'               " FROM TRADEMARK,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress,(" & strSQLE(0) & ") " & _
'               "WHERE tm10=na01(+) and " & IIf(strAppNo = strAppNo1, "TM79='" & strAppNo, "TM79>='" & strAppNo & "' and TM79<='" & strAppNo1) & "' and tm01 in (" & GetAddStr(Str02) & ") AND TM04='00' AND SUBSTR(TM23,1,8) = c1.CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1)) = c1.CU02(+) " & _
'               " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
'               " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) " & _
'               " and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) " & _
'               " and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) " & strSqlEW(0) & strSQL1
'      strSql = strSql + " union SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','')||Nvl(EState,'') AS 本所案號 ,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort,DECODE(TM123,NULL,C4.CU01||C4.CU127,C1.CU01||TM123) CNT" & SeColTM & _
'               " FROM TRADEMARK,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress,(" & strSQLE(0) & ") " & _
'               "WHERE tm10=na01(+) and " & IIf(strAppNo = strAppNo1, "TM80='" & strAppNo, "TM80>='" & strAppNo & "' AND TM80<='" & strAppNo1) & "' and tm01 in (" & GetAddStr(Str02) & ") AND TM04='00' AND SUBSTR(TM23,1,8) = c1.CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1)) = c1.CU02(+) " & _
'               " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
'               " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) " & _
'               " and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) " & _
'               " and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) " & strSqlEW(0) & strSQL1
'      strSql = strSql + " union SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','')||Nvl(EState,'') AS 本所案號 ,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort,DECODE(TM123,NULL,C5.CU01||C5.CU127,C1.CU01||TM123) CNT" & SeColTM & _
'               " FROM TRADEMARK,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress,(" & strSQLE(0) & ") " & _
'               "WHERE tm10=na01(+) and " & IIf(strAppNo = strAppNo1, "TM81='" & strAppNo, "TM81>='" & strAppNo & "' and TM81<='" & strAppNo1) & "' and tm01 in (" & GetAddStr(Str02) & ") AND TM04='00' AND SUBSTR(TM23,1,8) = c1.CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1)) = c1.CU02(+) " & _
'               " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
'               " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) " & _
'               " and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) " & _
'               " and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) " & strSqlEW(0) & strSQL1
'
'      'Modify By Sindy 2014/7/7 +||DECODE(PA165,'Y','＃','')
'      'Modified by Lydia 2019/12/26 增加欄位SeColPA
'      'Modify by Amy 2022/12/14 +e符號 Nvl(EState,'')
'      strSql = strSql + " union select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','')||Nvl(EState,'') AS 本所案號 ,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'               "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort,DECODE(PA149,NULL,C1.CU01||C1.CU127,C1.CU01||PA149) CNT" & SeColPA & _
'               " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress,(" & strSQLE(1) & ") " & _
'               " WHERE pa09=na01(+) and " & IIf(strAppNo = strAppNo1, "PA26='" & strAppNo, "PA26>='" & strAppNo & "' and PA26<='" & strAppNo1) & "' and pa01 in (" & GetAddStr(Str02) & ") and pa04='00' and substr(pa26,1,8)=c1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=c1.cu02(+) " & _
'               " and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'               " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) " & _
'               " and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) " & _
'               " and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) " & strSqlEW(1) & strSQL1
'      strSql = strSql + " union select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','')||Nvl(EState,'') AS 本所案號 ,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'               "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort,DECODE(PA149,NULL,C2.CU01||C2.CU127,C1.CU01||PA149) CNT" & SeColPA & _
'               " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress,(" & strSQLE(1) & ") " & _
'               " WHERE pa09=na01(+) and " & IIf(strAppNo = strAppNo1, "PA27='" & strAppNo, "PA27>='" & strAppNo & "' and PA27<='" & strAppNo1) & "' and pa01 in (" & GetAddStr(Str02) & ") and pa04='00' and substr(pa26,1,8)=c1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=c1.cu02(+) " & _
'               " and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'               " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) " & _
'               " and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) " & _
'               " and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) " & strSqlEW(1) & strSQL1
'      strSql = strSql + " union select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','')||Nvl(EState,'') AS 本所案號 ,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'               "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort,DECODE(PA149,NULL,C3.CU01||C3.CU127,C1.CU01||PA149) CNT" & SeColPA & _
'               " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress,(" & strSQLE(1) & ") " & _
'               " WHERE pa09=na01(+) and " & IIf(strAppNo = strAppNo1, "PA28='" & strAppNo, "PA28>='" & strAppNo & "' and PA28<='" & strAppNo1) & "' and pa01 in (" & GetAddStr(Str02) & ") and pa04='00' and substr(pa26,1,8)=c1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=c1.cu02(+) " & _
'               " and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'               " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) " & _
'               " and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) " & _
'               " and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) " & strSqlEW(1) & strSQL1
'      strSql = strSql + " union select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','')||Nvl(EState,'') AS 本所案號 ,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'               "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort,DECODE(PA149,NULL,C4.CU01||C4.CU127,C1.CU01||PA149) CNT" & SeColPA & _
'               " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress,(" & strSQLE(1) & ") " & _
'               " WHERE pa09=na01(+) and " & IIf(strAppNo = strAppNo1, "PA29='" & strAppNo, "PA29>='" & strAppNo & "' and PA29<='" & strAppNo1) & "' and pa01 in (" & GetAddStr(Str02) & ") and pa04='00' and substr(pa26,1,8)=c1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=c1.cu02(+) " & _
'               " and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'               " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) " & _
'               " and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) " & _
'               " and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and Pa01=cp01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) " & strSqlEW(1) & strSQL1
'      strSql = strSql + " union select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','')||Nvl(EState,'') AS 本所案號 ,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'               "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort,DECODE(PA149,NULL,C5.CU01||C5.CU127,C1.CU01||PA149) CNT" & SeColPA & _
'               " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress,(" & strSQLE(1) & ") " & _
'               " WHERE pa09=na01(+) and " & IIf(strAppNo = strAppNo1, "PA30='" & strAppNo, "PA30>='" & strAppNo & "' and PA30<='" & strAppNo1) & "' and pa01 in (" & GetAddStr(Str02) & ") and pa04='00' and substr(pa26,1,8)=c1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=c1.cu02(+) " & _
'               " and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'               " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) " & _
'               " and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) " & _
'               " and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) " & strSqlEW(1) & strSQL1
'
'      'Modified by Lydia 2019/12/26 增加欄位SeColSP
'      'Modify by Amy 2020/02/05 +SP73 商品類別
'      strSql = strSql + " union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'               "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort,DECODE(SP78,NULL,C1.CU01||C1.CU127,C1.CU01||SP78) CNT" & SeColSP & _
'               " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,customer c4,customer c5,caseprogress " & _
'               " WHERE sp09=na01(+) and " & IIf(strAppNo = strAppNo1, "sp08='" & strAppNo, "SP08>='" & strAppNo & "' and SP08<='" & strAppNo1) & "' and sp01 in (" & GetAddStr(Str02) & ") AND SP04='00' AND SUBSTR(SP08,1,8)=C1.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C1.CU02(+) " & _
'               " AND SUBSTr(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) " & _
'               " AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'               " and substr(sp65,1,8)=c4.cu01(+) and decode(substr(sp65,9,1),null,'0',substr(sp65,9,1))=c4.cu02(+) " & _
'               " and substr(sp66,1,8)=c5.cu01(+) and decode(substr(sp66,9,1),null,'0',substr(sp66,9,1))=c5.cu02(+) and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) " & strSQL1
'      strSql = strSql + " union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'               "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort,DECODE(SP78,NULL,C2.CU01||C2.CU127,C1.CU01||SP78) CNT" & SeColSP & _
'               " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,customer c4,customer c5,caseprogress " & _
'               " WHERE sp09=na01(+) and " & IIf(strAppNo = strAppNo1, "sp58='" & strAppNo, "SP58>='" & strAppNo & "' and SP58<='" & strAppNo1) & "' and sp01 in (" & GetAddStr(Str02) & ") AND SP04='00' AND SUBSTR(SP08,1,8)=C1.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C1.CU02(+) " & _
'               " AND SUBSTr(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) " & _
'               " AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'               " and substr(sp65,1,8)=c4.cu01(+) and decode(substr(sp65,9,1),null,'0',substr(sp65,9,1))=c4.cu02(+) " & _
'               " and substr(sp66,1,8)=c5.cu01(+) and decode(substr(sp66,9,1),null,'0',substr(sp66,9,1))=c5.cu02(+) and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) " & strSQL1
'      strSql = strSql + " union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'               "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort,DECODE(SP78,NULL,C3.CU01||C3.CU127,C1.CU01||SP78) CNT" & SeColSP & _
'               " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,customer c4,customer c5,caseprogress " & _
'               " WHERE sp09=na01(+) and " & IIf(strAppNo = strAppNo1, "sp59='" & strAppNo, "SP59>='" & strAppNo & "' and SP59<='" & strAppNo1) & "' and sp01 in (" & GetAddStr(Str02) & ") AND SP04='00' AND SUBSTR(SP08,1,8)=C1.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C1.CU02(+) " & _
'               " AND SUBSTr(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) " & _
'               " AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'               " and substr(sp65,1,8)=c4.cu01(+) and decode(substr(sp65,9,1),null,'0',substr(sp65,9,1))=c4.cu02(+) " & _
'               " and substr(sp66,1,8)=c5.cu01(+) and decode(substr(sp66,9,1),null,'0',substr(sp66,9,1))=c5.cu02(+) and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) " & strSQL1
'      strSql = strSql + " union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'               "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort,DECODE(SP78,NULL,C4.CU01||C4.CU127,C1.CU01||SP78) CNT" & SeColSP & _
'               " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,customer c4,customer c5,caseprogress " & _
'               " WHERE sp09=na01(+) and " & IIf(strAppNo = strAppNo1, "sp65='" & strAppNo, "SP65>='" & strAppNo & "' and SP65<='" & strAppNo1) & "' and sp01 in (" & GetAddStr(Str02) & ") AND SP04='00' AND SUBSTR(SP08,1,8)=C1.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C1.CU02(+) " & _
'               " AND SUBSTr(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) " & _
'               " AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'               " and substr(sp65,1,8)=c4.cu01(+) and decode(substr(sp65,9,1),null,'0',substr(sp65,9,1))=c4.cu02(+) " & _
'               " and substr(sp66,1,8)=c5.cu01(+) and decode(substr(sp66,9,1),null,'0',substr(sp66,9,1))=c5.cu02(+) and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) " & strSQL1
'      strSql = strSql + " union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'               "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort,DECODE(SP78,NULL,C5.CU01||C5.CU127,C1.CU01||SP78) CNT" & SeColSP & _
'               " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,customer c4,customer c5,caseprogress " & _
'               " WHERE sp09=na01(+) and " & IIf(strAppNo = strAppNo1, "sp66='" & strAppNo, "SP66>='" & strAppNo & "' and SP66<='" & strAppNo1) & "' and sp01 in (" & GetAddStr(Str02) & ") AND SP04='00' AND SUBSTR(SP08,1,8)=C1.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C1.CU02(+) " & _
'               " AND SUBSTr(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) " & _
'               " AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'               " and substr(sp65,1,8)=c4.cu01(+) and decode(substr(sp65,9,1),null,'0',substr(sp65,9,1))=c4.cu02(+) " & _
'               " and substr(sp66,1,8)=c5.cu01(+) and decode(substr(sp66,9,1),null,'0',substr(sp66,9,1))=c5.cu02(+) and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) " & strSQL1
'      'end 2020/02/05
'
'      'Modify By Sindy 2011/1/20 +LC43,LC44,LC45,LC46
'      'Modify by Amy 2018/09/17 不加專案服務案 lc52=Y 顯示 案件進度+案件性質
'      'Modified by Lydia 2019/12/26 增加欄位SeColLC
'      strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(c1.CU04,NVL(c1.CU05||c1.CU88||c1.CU89||c1.CU90,c1.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2, NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3, NVL(c4.CU04,NVL(c4.CU05||c4.CU88||c4.CU89||c4.CU90,c4.CU06)) AS 申請人4, NVL(c5.CU04,NVL(c5.CU05||c5.CU88||c5.CU89||c5.CU90,c5.CU06)) AS 申請人5, LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort" & _
'               ",DECODE(LC42,NULL,c1.CU01||c1.CU127,c1.CU01||LC42) CNT" & SeColLC & _
'               " FROM LAWCASE,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress,CasePropertyMap " & _
'               " WHERE lc15=na01(+) and " & IIf(strAppNo = strAppNo1, "LC11='" & strAppNo, "LC11>='" & strAppNo & "' and LC11<='" & strAppNo1) & "' and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00' AND SUBSTR(LC11,1,8)=c1.CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = c1.CU02(+) " & _
'               " and SUBSTr(lc43,1,8)=C2.CU01(+) AND DECODE(SUBSTR(lc43,9,1),NULL,'0',SUBSTR(lc43,9,1))=C2.CU02(+) " & _
'               " and SUBSTr(lc44,1,8)=C3.CU01(+) AND DECODE(SUBSTR(lc44,9,1),NULL,'0',SUBSTR(lc44,9,1))=C3.CU02(+) " & _
'               " and SUBSTr(lc45,1,8)=C4.CU01(+) AND DECODE(SUBSTR(lc45,9,1),NULL,'0',SUBSTR(lc45,9,1))=C4.CU02(+) " & _
'               " and SUBSTr(lc46,1,8)=C5.CU01(+) AND DECODE(SUBSTR(lc46,9,1),NULL,'0',SUBSTR(lc46,9,1))=C5.CU02(+) and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+)  and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1
'      strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(c1.CU04,NVL(c1.CU05||c1.CU88||c1.CU89||c1.CU90,c1.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2, NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3, NVL(c4.CU04,NVL(c4.CU05||c4.CU88||c4.CU89||c4.CU90,c4.CU06)) AS 申請人4, NVL(c5.CU04,NVL(c5.CU05||c5.CU88||c5.CU89||c5.CU90,c5.CU06)) AS 申請人5, LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort" & _
'               ",DECODE(LC42,NULL,c1.CU01||c1.CU127,c1.CU01||LC42) CNT" & SeColLC & _
'               " FROM LAWCASE,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress,CasePropertyMap " & _
'               " WHERE lc15=na01(+) and " & IIf(strAppNo = strAppNo1, "LC43='" & strAppNo, "LC43>='" & strAppNo & "' and LC43<='" & strAppNo1) & "' and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00' AND SUBSTR(LC11,1,8)=c1.CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = c1.CU02(+) " & _
'               " and SUBSTr(lc43,1,8)=C2.CU01(+) AND DECODE(SUBSTR(lc43,9,1),NULL,'0',SUBSTR(lc43,9,1))=C2.CU02(+) " & _
'               " and SUBSTr(lc44,1,8)=C3.CU01(+) AND DECODE(SUBSTR(lc44,9,1),NULL,'0',SUBSTR(lc44,9,1))=C3.CU02(+) " & _
'               " and SUBSTr(lc45,1,8)=C4.CU01(+) AND DECODE(SUBSTR(lc45,9,1),NULL,'0',SUBSTR(lc45,9,1))=C4.CU02(+) " & _
'               " and SUBSTr(lc46,1,8)=C5.CU01(+) AND DECODE(SUBSTR(lc46,9,1),NULL,'0',SUBSTR(lc46,9,1))=C5.CU02(+) and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1
'      strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(c1.CU04,NVL(c1.CU05||c1.CU88||c1.CU89||c1.CU90,c1.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2, NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3, NVL(c4.CU04,NVL(c4.CU05||c4.CU88||c4.CU89||c4.CU90,c4.CU06)) AS 申請人4, NVL(c5.CU04,NVL(c5.CU05||c5.CU88||c5.CU89||c5.CU90,c5.CU06)) AS 申請人5, LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort" & _
'               ",DECODE(LC42,NULL,c1.CU01||c1.CU127,c1.CU01||LC42) CNT" & SeColLC & _
'               " FROM LAWCASE,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress,CasePropertyMap " & _
'               " WHERE lc15=na01(+) and " & IIf(strAppNo = strAppNo1, "LC44='" & strAppNo, "LC44>='" & strAppNo & "' and LC44<='" & strAppNo1) & "' and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00' AND SUBSTR(LC11,1,8)=c1.CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = c1.CU02(+) " & _
'               " and SUBSTr(lc43,1,8)=C2.CU01(+) AND DECODE(SUBSTR(lc43,9,1),NULL,'0',SUBSTR(lc43,9,1))=C2.CU02(+) " & _
'               " and SUBSTr(lc44,1,8)=C3.CU01(+) AND DECODE(SUBSTR(lc44,9,1),NULL,'0',SUBSTR(lc44,9,1))=C3.CU02(+) " & _
'               " and SUBSTr(lc45,1,8)=C4.CU01(+) AND DECODE(SUBSTR(lc45,9,1),NULL,'0',SUBSTR(lc45,9,1))=C4.CU02(+) " & _
'               " and SUBSTr(lc46,1,8)=C5.CU01(+) AND DECODE(SUBSTR(lc46,9,1),NULL,'0',SUBSTR(lc46,9,1))=C5.CU02(+) and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1
'      strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(c1.CU04,NVL(c1.CU05||c1.CU88||c1.CU89||c1.CU90,c1.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2, NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3, NVL(c4.CU04,NVL(c4.CU05||c4.CU88||c4.CU89||c4.CU90,c4.CU06)) AS 申請人4, NVL(c5.CU04,NVL(c5.CU05||c5.CU88||c5.CU89||c5.CU90,c5.CU06)) AS 申請人5, LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort" & _
'               ",DECODE(LC42,NULL,c1.CU01||c1.CU127,c1.CU01||LC42) CNT" & SeColLC & _
'               " FROM LAWCASE,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress,CasePropertyMap " & _
'               " WHERE lc15=na01(+) and " & IIf(strAppNo = strAppNo1, "LC45='" & strAppNo, "LC45>='" & strAppNo & "' and LC45<='" & strAppNo1) & "' and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00' AND SUBSTR(LC11,1,8)=c1.CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = c1.CU02(+) " & _
'               " and SUBSTr(lc43,1,8)=C2.CU01(+) AND DECODE(SUBSTR(lc43,9,1),NULL,'0',SUBSTR(lc43,9,1))=C2.CU02(+) " & _
'               " and SUBSTr(lc44,1,8)=C3.CU01(+) AND DECODE(SUBSTR(lc44,9,1),NULL,'0',SUBSTR(lc44,9,1))=C3.CU02(+) " & _
'               " and SUBSTr(lc45,1,8)=C4.CU01(+) AND DECODE(SUBSTR(lc45,9,1),NULL,'0',SUBSTR(lc45,9,1))=C4.CU02(+) " & _
'               " and SUBSTr(lc46,1,8)=C5.CU01(+) AND DECODE(SUBSTR(lc46,9,1),NULL,'0',SUBSTR(lc46,9,1))=C5.CU02(+) and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1
'      strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(c1.CU04,NVL(c1.CU05||c1.CU88||c1.CU89||c1.CU90,c1.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2, NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3, NVL(c4.CU04,NVL(c4.CU05||c4.CU88||c4.CU89||c4.CU90,c4.CU06)) AS 申請人4, NVL(c5.CU04,NVL(c5.CU05||c5.CU88||c5.CU89||c5.CU90,c5.CU06)) AS 申請人5, LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort" & _
'               ",DECODE(LC42,NULL,c1.CU01||c1.CU127,c1.CU01||LC42) CNT" & SeColLC & _
'               " FROM LAWCASE,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress,CasePropertyMap " & _
'               " WHERE lc15=na01(+) and " & IIf(strAppNo = strAppNo1, "LC46='" & strAppNo, "LC46>='" & strAppNo & "' and LC46<='" & strAppNo1) & "' and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00' AND SUBSTR(LC11,1,8)=c1.CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = c1.CU02(+) " & _
'               " and SUBSTr(lc43,1,8)=C2.CU01(+) AND DECODE(SUBSTR(lc43,9,1),NULL,'0',SUBSTR(lc43,9,1))=C2.CU02(+) " & _
'               " and SUBSTr(lc44,1,8)=C3.CU01(+) AND DECODE(SUBSTR(lc44,9,1),NULL,'0',SUBSTR(lc44,9,1))=C3.CU02(+) " & _
'               " and SUBSTr(lc45,1,8)=C4.CU01(+) AND DECODE(SUBSTR(lc45,9,1),NULL,'0',SUBSTR(lc45,9,1))=C4.CU02(+) " & _
'               " and SUBSTr(lc46,1,8)=C5.CU01(+) AND DECODE(SUBSTR(lc46,9,1),NULL,'0',SUBSTR(lc46,9,1))=C5.CU02(+) and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1
'
'      'Modify By Sindy 2011/1/20 +HC24,HC25,HC26,HC27
'      'Modified by Lydia 2019/12/26 增加欄位SeColHC
'      strSql = strSql + " union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,'台灣' AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(c1.CU04,NVL(c1.CU05||c1.CU88||c1.CU89||c1.CU90,c1.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2, NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3, NVL(c4.CU04,NVL(c4.CU05||c4.CU88||c4.CU89||c4.CU90,c4.CU06)) AS 申請人4, NVL(c5.CU04,NVL(c5.CU05||c5.CU88||c5.CU89||c5.CU90,c5.CU06)) AS 申請人5, HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort" & _
'               ",DECODE(HC23,NULL,c1.CU01||c1.CU127,c1.CU01||HC23) CNT" & SeColHC & _
'               " FROM HIRECASE,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress " & _
'               " WHERE " & IIf(strAppNo = strAppNo1, "HC05='" & strAppNo, "HC05>='" & strAppNo & "' and HC05<='" & strAppNo1) & "' and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00' AND SUBSTR(HC05,1,8)=c1.CU01(+) AND DECODE(SUBSTR(HC05,9,1),NULL,'0',SUBSTR(HC05,9,1))=c1.CU02(+) " & _
'               " and SUBSTr(hc24,1,8)=C2.CU01(+) AND DECODE(SUBSTR(hc24,9,1),NULL,'0',SUBSTR(hc24,9,1))=C2.CU02(+) " & _
'               " and SUBSTr(hc25,1,8)=C3.CU01(+) AND DECODE(SUBSTR(hc25,9,1),NULL,'0',SUBSTR(hc25,9,1))=C3.CU02(+) " & _
'               " and SUBSTr(hc26,1,8)=C4.CU01(+) AND DECODE(SUBSTR(hc26,9,1),NULL,'0',SUBSTR(hc26,9,1))=C4.CU02(+) " & _
'               " and SUBSTr(hc27,1,8)=C5.CU01(+) AND DECODE(SUBSTR(hc27,9,1),NULL,'0',SUBSTR(hc27,9,1))=C5.CU02(+) and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) " & strSQL1
'      strSql = strSql + " union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,'台灣' AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(c1.CU04,NVL(c1.CU05||c1.CU88||c1.CU89||c1.CU90,c1.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2, NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3, NVL(c4.CU04,NVL(c4.CU05||c4.CU88||c4.CU89||c4.CU90,c4.CU06)) AS 申請人4, NVL(c5.CU04,NVL(c5.CU05||c5.CU88||c5.CU89||c5.CU90,c5.CU06)) AS 申請人5, HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort" & _
'               ",DECODE(HC23,NULL,c1.CU01||c1.CU127,c1.CU01||HC23) CNT" & SeColHC & _
'               " FROM HIRECASE,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress " & _
'               " WHERE " & IIf(strAppNo = strAppNo1, "hc24='" & strAppNo, "hc24>='" & strAppNo & "' and hc24<='" & strAppNo1) & "' and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00' AND SUBSTR(HC05,1,8)=c1.CU01(+) AND DECODE(SUBSTR(HC05,9,1),NULL,'0',SUBSTR(HC05,9,1))=c1.CU02(+) " & _
'               " and SUBSTr(hc24,1,8)=C2.CU01(+) AND DECODE(SUBSTR(hc24,9,1),NULL,'0',SUBSTR(hc24,9,1))=C2.CU02(+) " & _
'               " and SUBSTr(hc25,1,8)=C3.CU01(+) AND DECODE(SUBSTR(hc25,9,1),NULL,'0',SUBSTR(hc25,9,1))=C3.CU02(+) " & _
'               " and SUBSTr(hc26,1,8)=C4.CU01(+) AND DECODE(SUBSTR(hc26,9,1),NULL,'0',SUBSTR(hc26,9,1))=C4.CU02(+) " & _
'               " and SUBSTr(hc27,1,8)=C5.CU01(+) AND DECODE(SUBSTR(hc27,9,1),NULL,'0',SUBSTR(hc27,9,1))=C5.CU02(+) and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) " & strSQL1
'      strSql = strSql + " union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,'台灣' AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(c1.CU04,NVL(c1.CU05||c1.CU88||c1.CU89||c1.CU90,c1.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2, NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3, NVL(c4.CU04,NVL(c4.CU05||c4.CU88||c4.CU89||c4.CU90,c4.CU06)) AS 申請人4, NVL(c5.CU04,NVL(c5.CU05||c5.CU88||c5.CU89||c5.CU90,c5.CU06)) AS 申請人5, HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort" & _
'               ",DECODE(HC23,NULL,c1.CU01||c1.CU127,c1.CU01||HC23) CNT" & SeColHC & _
'               " FROM HIRECASE,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress " & _
'               " WHERE " & IIf(strAppNo = strAppNo1, "hc25='" & strAppNo, "hc25>='" & strAppNo & "' and hc25<='" & strAppNo1) & "' and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00' AND SUBSTR(HC05,1,8)=c1.CU01(+) AND DECODE(SUBSTR(HC05,9,1),NULL,'0',SUBSTR(HC05,9,1))=c1.CU02(+) " & _
'               " and SUBSTr(hc24,1,8)=C2.CU01(+) AND DECODE(SUBSTR(hc24,9,1),NULL,'0',SUBSTR(hc24,9,1))=C2.CU02(+) " & _
'               " and SUBSTr(hc25,1,8)=C3.CU01(+) AND DECODE(SUBSTR(hc25,9,1),NULL,'0',SUBSTR(hc25,9,1))=C3.CU02(+) " & _
'               " and SUBSTr(hc26,1,8)=C4.CU01(+) AND DECODE(SUBSTR(hc26,9,1),NULL,'0',SUBSTR(hc26,9,1))=C4.CU02(+) " & _
'               " and SUBSTr(hc27,1,8)=C5.CU01(+) AND DECODE(SUBSTR(hc27,9,1),NULL,'0',SUBSTR(hc27,9,1))=C5.CU02(+) and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) " & strSQL1
'      strSql = strSql + " union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,'台灣' AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(c1.CU04,NVL(c1.CU05||c1.CU88||c1.CU89||c1.CU90,c1.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2, NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3, NVL(c4.CU04,NVL(c4.CU05||c4.CU88||c4.CU89||c4.CU90,c4.CU06)) AS 申請人4, NVL(c5.CU04,NVL(c5.CU05||c5.CU88||c5.CU89||c5.CU90,c5.CU06)) AS 申請人5, HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort" & _
'               ",DECODE(HC23,NULL,c1.CU01||c1.CU127,c1.CU01||HC23) CNT" & SeColHC & _
'               " FROM HIRECASE,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress " & _
'               " WHERE " & IIf(strAppNo = strAppNo1, "hc26='" & strAppNo, "hc26>='" & strAppNo & "' and hc26<='" & strAppNo1) & "' and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00' AND SUBSTR(HC05,1,8)=c1.CU01(+) AND DECODE(SUBSTR(HC05,9,1),NULL,'0',SUBSTR(HC05,9,1))=c1.CU02(+) " & _
'               " and SUBSTr(hc24,1,8)=C2.CU01(+) AND DECODE(SUBSTR(hc24,9,1),NULL,'0',SUBSTR(hc24,9,1))=C2.CU02(+) " & _
'               " and SUBSTr(hc25,1,8)=C3.CU01(+) AND DECODE(SUBSTR(hc25,9,1),NULL,'0',SUBSTR(hc25,9,1))=C3.CU02(+) " & _
'               " and SUBSTr(hc26,1,8)=C4.CU01(+) AND DECODE(SUBSTR(hc26,9,1),NULL,'0',SUBSTR(hc26,9,1))=C4.CU02(+) " & _
'               " and SUBSTr(hc27,1,8)=C5.CU01(+) AND DECODE(SUBSTR(hc27,9,1),NULL,'0',SUBSTR(hc27,9,1))=C5.CU02(+) and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) " & strSQL1
'      strSql = strSql + " union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,'台灣' AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(c1.CU04,NVL(c1.CU05||c1.CU88||c1.CU89||c1.CU90,c1.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2, NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3, NVL(c4.CU04,NVL(c4.CU05||c4.CU88||c4.CU89||c4.CU90,c4.CU06)) AS 申請人4, NVL(c5.CU04,NVL(c5.CU05||c5.CU88||c5.CU89||c5.CU90,c5.CU06)) AS 申請人5, HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort" & _
'               ",DECODE(HC23,NULL,c1.CU01||c1.CU127,c1.CU01||HC23) CNT" & SeColHC & _
'               " FROM HIRECASE,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress " & _
'               " WHERE " & IIf(strAppNo = strAppNo1, "hc27='" & strAppNo, "hc27>='" & strAppNo & "' and hc27<='" & strAppNo1) & "' and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00' AND SUBSTR(HC05,1,8)=c1.CU01(+) AND DECODE(SUBSTR(HC05,9,1),NULL,'0',SUBSTR(HC05,9,1))=c1.CU02(+) " & _
'               " and SUBSTr(hc24,1,8)=C2.CU01(+) AND DECODE(SUBSTR(hc24,9,1),NULL,'0',SUBSTR(hc24,9,1))=C2.CU02(+) " & _
'               " and SUBSTr(hc25,1,8)=C3.CU01(+) AND DECODE(SUBSTR(hc25,9,1),NULL,'0',SUBSTR(hc25,9,1))=C3.CU02(+) " & _
'               " and SUBSTr(hc26,1,8)=C4.CU01(+) AND DECODE(SUBSTR(hc26,9,1),NULL,'0',SUBSTR(hc26,9,1))=C4.CU02(+) " & _
'               " and SUBSTr(hc27,1,8)=C5.CU01(+) AND DECODE(SUBSTR(hc27,9,1),NULL,'0',SUBSTR(hc27,9,1))=C5.CU02(+) and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) " & strSQL1
'
'      '2009/3/6 MODIFY BY SONIA 商標之申請人1抓錯,原抓CP55串之C1的CUSTOMER
'      '加考慮案件進度檔CP55 , cp56, CP72欄位
'      'Modified by Morgan 2019/1/17 查詢接洽人案件時,非申請人的接洽人編號不必抓個案
'      'Modified by Lydia 2019/12/26 增加欄位SeColTM
'      'Modify by Amy 2022/11/14 +strWhereCP 以X編號抓進度檔的CP55/CP56/CP89~CP96的資料時,要加入條件(CP158>0 OR CP159=0)
'      'Modify by Amy 2022/12/14 +e符號 Nvl(EState,'')
'      strSql = strSql + " union SELECT ' ' AS V,'△'||decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','')||Nvl(EState,'') AS 本所案號 , DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'         ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort,C1.CU01||C1.CU127 CNT" & SeColTM & _
'         " FROM TRADEMARK,nation,CUSTOMER C1, CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress,(" & strSQLE(0) & ") " & _
'         "WHERE tm10=na01(+) and " & IIf(strAppNo = strAppNo1, "CP55='" & strAppNo, "CP55>='" & strAppNo & "' and CP55<='" & strAppNo1) & "' " & strWhereCP & " And (TM23<>'" & Me.Tag & "' Or TM23 Is Null) And (TM78<>'" & Me.Tag & "' Or TM78 Is Null) And (TM79<>'" & Me.Tag & "' Or TM79 Is Null) And (TM80<>'" & Me.Tag & "' Or TM80 Is Null) And (TM81<>'" & Me.Tag & "' Or TM81 Is Null) and tm01 in (" & GetAddStr(Str02) & ") AND TM04='00' AND SUBSTR(cp55,1,8)=C1.CU01(+) AND DECODE(SUBSTR(cp55,9,1),NULL,'0',SUBSTR(cp55,9,1))=C1.CU02(+)" & _
'         " and substr(tm23,1,8)=c6.cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=c6.cu02(+) " & _
'         " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
'         " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) " & _
'         " and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) " & _
'         " and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+)  and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & strSqlEW(0) & strSQL1
'      strSql = strSql + " union SELECT ' ' AS V,'△'||decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','')||Nvl(EState,'') AS 本所案號 , DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'         ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort,C1.CU01||C1.CU127 CNT" & SeColTM & _
'         " FROM TRADEMARK,nation,CUSTOMER C1, CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress,(" & strSQLE(0) & ") " & _
'         "WHERE tm10=na01(+) and " & IIf(strAppNo = strAppNo1, "CP56='" & strAppNo, "CP56>='" & strAppNo & "' and CP56<='" & strAppNo1) & "' " & strWhereCP & " And (TM23<>'" & Me.Tag & "' Or TM23 Is Null) And (TM78<>'" & Me.Tag & "' Or TM78 Is Null) And (TM79<>'" & Me.Tag & "' Or TM79 Is Null) And (TM80<>'" & Me.Tag & "' Or TM80 Is Null) And (TM81<>'" & Me.Tag & "' Or TM81 Is Null) and tm01 in (" & GetAddStr(Str02) & ") AND TM04='00' AND SUBSTR(cp56,1,8)=C1.CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=C1.CU02(+)" & _
'         " and substr(tm23,1,8)=c6.cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=c6.cu02(+) " & _
'         " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
'         " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) " & _
'         " and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) " & _
'         " and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & strSqlEW(0) & strSQL1
'      strSql = strSql + " union SELECT ' ' AS V,'△'||decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','')||Nvl(EState,'') AS 本所案號 , DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'         ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort,C1.CU01||C1.CU127 CNT" & SeColTM & _
'         " FROM TRADEMARK,nation,CUSTOMER C1, CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress,(" & strSQLE(0) & ") " & _
'         "WHERE tm10=na01(+) and " & IIf(strAppNo = strAppNo1, "CP72='" & strAppNo, "CP72>='" & strAppNo & "' and CP72<='" & strAppNo1) & "' And (TM23<>'" & Me.Tag & "' Or TM23 Is Null) And (TM78<>'" & Me.Tag & "' Or TM78 Is Null) And (TM79<>'" & Me.Tag & "' Or TM79 Is Null) And (TM80<>'" & Me.Tag & "' Or TM80 Is Null) And (TM81<>'" & Me.Tag & "' Or TM81 Is Null) and tm01 in (" & GetAddStr(Str02) & ") AND TM04='00' AND SUBSTR(cp72,1,8)=C1.CU01(+) AND DECODE(SUBSTR(cp72,9,1),NULL,'0',SUBSTR(cp72,9,1))=C1.CU02(+)" & _
'         " and substr(tm23,1,8)=c6.cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=c6.cu02(+) " & _
'         " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
'         " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) " & _
'         " and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) " & _
'         " and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & strSqlEW(0) & strSQL1
'      '2008/7/31 add by sonia 加案件進度檔 cp89~cp96
'      strSql = strSql + " union SELECT ' ' AS V,'△'||decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','')||Nvl(EState,'') AS 本所案號 , DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'         ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort,C1.CU01||C1.CU127 CNT" & SeColTM & _
'         " FROM TRADEMARK,nation,CUSTOMER C1, CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress,(" & strSQLE(0) & ") " & _
'         "WHERE tm10=na01(+) and " & IIf(strAppNo = strAppNo1, "CP89='" & strAppNo, "CP89>='" & strAppNo & "' and CP89<='" & strAppNo1) & "' " & strWhereCP & " And (TM23<>'" & Me.Tag & "' Or TM23 Is Null) And (TM78<>'" & Me.Tag & "' Or TM78 Is Null) And (TM79<>'" & Me.Tag & "' Or TM79 Is Null) And (TM80<>'" & Me.Tag & "' Or TM80 Is Null) And (TM81<>'" & Me.Tag & "' Or TM81 Is Null) and tm01 in (" & GetAddStr(Str02) & ") AND TM04='00' AND SUBSTR(cp89,1,8)=C1.CU01(+) AND DECODE(SUBSTR(cp89,9,1),NULL,'0',SUBSTR(cp89,9,1))=C1.CU02(+)" & _
'         " and substr(tm23,1,8)=c6.cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=c6.cu02(+) " & _
'         " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
'         " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) " & _
'         " and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) " & _
'         " and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & strSqlEW(0) & strSQL1
'      strSql = strSql + " union SELECT ' ' AS V,'△'||decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','')||Nvl(EState,'') AS 本所案號 , DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'         ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort,C1.CU01||C1.CU127 CNT" & SeColTM & _
'         " FROM TRADEMARK,nation,CUSTOMER C1, CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress,(" & strSQLE(0) & ") " & _
'         "WHERE tm10=na01(+) and " & IIf(strAppNo = strAppNo1, "CP90='" & strAppNo, "CP90>='" & strAppNo & "' and CP90<='" & strAppNo1) & "' " & strWhereCP & " And (TM23<>'" & Me.Tag & "' Or TM23 Is Null) And (TM78<>'" & Me.Tag & "' Or TM78 Is Null) And (TM79<>'" & Me.Tag & "' Or TM79 Is Null) And (TM80<>'" & Me.Tag & "' Or TM80 Is Null) And (TM81<>'" & Me.Tag & "' Or TM81 Is Null) and tm01 in (" & GetAddStr(Str02) & ") AND TM04='00' AND SUBSTR(cp90,1,8)=C1.CU01(+) AND DECODE(SUBSTR(cp90,9,1),NULL,'0',SUBSTR(cp90,9,1))=C1.CU02(+)" & _
'         " and substr(tm23,1,8)=c6.cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=c6.cu02(+) " & _
'         " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
'         " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) " & _
'         " and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) " & _
'         " and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & strSqlEW(0) & strSQL1
'      strSql = strSql + " union SELECT ' ' AS V,'△'||decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','')||Nvl(EState,'') AS 本所案號 , DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'         ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort,C1.CU01||C1.CU127 CNT" & SeColTM & _
'         " FROM TRADEMARK,nation,CUSTOMER C1, CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress,(" & strSQLE(0) & ") " & _
'         "WHERE tm10=na01(+) and " & IIf(strAppNo = strAppNo1, "CP91='" & strAppNo, "CP91>='" & strAppNo & "' and CP91<='" & strAppNo1) & "' " & strWhereCP & " And (TM23<>'" & Me.Tag & "' Or TM23 Is Null) And (TM78<>'" & Me.Tag & "' Or TM78 Is Null) And (TM79<>'" & Me.Tag & "' Or TM79 Is Null) And (TM80<>'" & Me.Tag & "' Or TM80 Is Null) And (TM81<>'" & Me.Tag & "' Or TM81 Is Null) and tm01 in (" & GetAddStr(Str02) & ") AND TM04='00' AND SUBSTR(cp91,1,8)=C1.CU01(+) AND DECODE(SUBSTR(cp91,9,1),NULL,'0',SUBSTR(cp91,9,1))=C1.CU02(+)" & _
'         " and substr(tm23,1,8)=c6.cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=c6.cu02(+) " & _
'         " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
'         " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) " & _
'         " and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) " & _
'         " and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & strSqlEW(0) & strSQL1
'      strSql = strSql + " union SELECT ' ' AS V,'△'||decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','')||Nvl(EState,'') AS 本所案號 , DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'         ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort,C1.CU01||C1.CU127 CNT" & SeColTM & _
'         " FROM TRADEMARK,nation,CUSTOMER C1, CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress,(" & strSQLE(0) & ") " & _
'         "WHERE tm10=na01(+) and " & IIf(strAppNo = strAppNo1, "CP92='" & strAppNo, "CP92>='" & strAppNo & "' and CP92<='" & strAppNo1) & "' " & strWhereCP & " And (TM23<>'" & Me.Tag & "' Or TM23 Is Null) And (TM78<>'" & Me.Tag & "' Or TM78 Is Null) And (TM79<>'" & Me.Tag & "' Or TM79 Is Null) And (TM80<>'" & Me.Tag & "' Or TM80 Is Null) And (TM81<>'" & Me.Tag & "' Or TM81 Is Null) and tm01 in (" & GetAddStr(Str02) & ") AND TM04='00' AND SUBSTR(cp92,1,8)=C1.CU01(+) AND DECODE(SUBSTR(cp92,9,1),NULL,'0',SUBSTR(cp92,9,1))=C1.CU02(+)" & _
'         " and substr(tm23,1,8)=c6.cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=c6.cu02(+) " & _
'         " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
'         " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) " & _
'         " and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) " & _
'         " and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & strSqlEW(0) & strSQL1
'      strSql = strSql + " union SELECT ' ' AS V,'△'||decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','')||Nvl(EState,'') AS 本所案號 , DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'         ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort,C1.CU01||C1.CU127 CNT" & SeColTM & _
'         " FROM TRADEMARK,nation,CUSTOMER C1, CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress,(" & strSQLE(0) & ") " & _
'         "WHERE tm10=na01(+) and " & IIf(strAppNo = strAppNo1, "CP93='" & strAppNo, "CP93>='" & strAppNo & "' and CP93<='" & strAppNo1) & "' " & strWhereCP & " And (TM23<>'" & Me.Tag & "' Or TM23 Is Null) And (TM78<>'" & Me.Tag & "' Or TM78 Is Null) And (TM79<>'" & Me.Tag & "' Or TM79 Is Null) And (TM80<>'" & Me.Tag & "' Or TM80 Is Null) And (TM81<>'" & Me.Tag & "' Or TM81 Is Null) and tm01 in (" & GetAddStr(Str02) & ") AND TM04='00' AND SUBSTR(cp93,1,8)=C1.CU01(+) AND DECODE(SUBSTR(cp93,9,1),NULL,'0',SUBSTR(cp93,9,1))=C1.CU02(+)" & _
'         " and substr(tm23,1,8)=c6.cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=c6.cu02(+) " & _
'         " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
'         " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) " & _
'         " and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) " & _
'         " and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & strSqlEW(0) & strSQL1
'      strSql = strSql + " union SELECT ' ' AS V,'△'||decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','')||Nvl(EState,'') AS 本所案號 , DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'         ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort,C1.CU01||C1.CU127 CNT" & SeColTM & _
'         " FROM TRADEMARK,nation,CUSTOMER C1, CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress,(" & strSQLE(0) & ") " & _
'         "WHERE tm10=na01(+) and " & IIf(strAppNo = strAppNo1, "CP94='" & strAppNo, "CP94>='" & strAppNo & "' and CP94<='" & strAppNo1) & "' " & strWhereCP & " And (TM23<>'" & Me.Tag & "' Or TM23 Is Null) And (TM78<>'" & Me.Tag & "' Or TM78 Is Null) And (TM79<>'" & Me.Tag & "' Or TM79 Is Null) And (TM80<>'" & Me.Tag & "' Or TM80 Is Null) And (TM81<>'" & Me.Tag & "' Or TM81 Is Null) and tm01 in (" & GetAddStr(Str02) & ") AND TM04='00' AND SUBSTR(cp94,1,8)=C1.CU01(+) AND DECODE(SUBSTR(cp94,9,1),NULL,'0',SUBSTR(cp94,9,1))=C1.CU02(+)" & _
'         " and substr(tm23,1,8)=c6.cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=c6.cu02(+) " & _
'         " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
'         " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) " & _
'         " and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) " & _
'         " and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & strSqlEW(0) & strSQL1
'      strSql = strSql + " union SELECT ' ' AS V,'△'||decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','')||Nvl(EState,'') AS 本所案號 , DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'         ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort,C1.CU01||C1.CU127 CNT" & SeColTM & _
'         " FROM TRADEMARK,nation,CUSTOMER C1, CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress,(" & strSQLE(0) & ") " & _
'         "WHERE tm10=na01(+) and " & IIf(strAppNo = strAppNo1, "CP95='" & strAppNo, "CP95>='" & strAppNo & "' and CP95<='" & strAppNo1) & "' " & strWhereCP & " And (TM23<>'" & Me.Tag & "' Or TM23 Is Null) And (TM78<>'" & Me.Tag & "' Or TM78 Is Null) And (TM79<>'" & Me.Tag & "' Or TM79 Is Null) And (TM80<>'" & Me.Tag & "' Or TM80 Is Null) And (TM81<>'" & Me.Tag & "' Or TM81 Is Null) and tm01 in (" & GetAddStr(Str02) & ") AND TM04='00' AND SUBSTR(cp95,1,8)=C1.CU01(+) AND DECODE(SUBSTR(cp95,9,1),NULL,'0',SUBSTR(cp95,9,1))=C1.CU02(+)" & _
'         " and substr(tm23,1,8)=c6.cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=c6.cu02(+) " & _
'         " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
'         " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) " & _
'         " and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) " & _
'         " and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+)  and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & strSqlEW(0) & strSQL1
'      strSql = strSql + " union SELECT ' ' AS V,'△'||decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','')||Nvl(EState,'') AS 本所案號 , DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'         ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort,C1.CU01||C1.CU127 CNT" & SeColTM & _
'         " FROM TRADEMARK,nation,CUSTOMER C1, CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress,(" & strSQLE(0) & ") " & _
'         "WHERE tm10=na01(+) and " & IIf(strAppNo = strAppNo1, "CP96='" & strAppNo, "CP96>='" & strAppNo & "' and CP96<='" & strAppNo1) & "' " & strWhereCP & " And (TM23<>'" & Me.Tag & "' Or TM23 Is Null) And (TM78<>'" & Me.Tag & "' Or TM78 Is Null) And (TM79<>'" & Me.Tag & "' Or TM79 Is Null) And (TM80<>'" & Me.Tag & "' Or TM80 Is Null) And (TM81<>'" & Me.Tag & "' Or TM81 Is Null) and tm01 in (" & GetAddStr(Str02) & ") AND TM04='00' AND SUBSTR(cp96,1,8)=C1.CU01(+) AND DECODE(SUBSTR(cp96,9,1),NULL,'0',SUBSTR(cp96,9,1))=C1.CU02(+)" & _
'         " and substr(tm23,1,8)=c6.cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=c6.cu02(+) " & _
'         " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
'         " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) " & _
'         " and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) " & _
'         " and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & strSqlEW(0) & strSQL1
'      '2008/7/31 end
'
'      'edit by nickc 2006/06/08 申請人1長度不要限制
'      'Modify By Sindy 2014/7/7 +||DECODE(PA165,'Y','＃','')
'      'Modified by Lydia 2019/12/26 增加欄位SeColPA
'      'Modify by Amy 2022/11/14 +strWhereCP 以X編號抓進度檔的CP55/CP56/CP89~CP96的資料時,要加入條件(CP158>0 OR CP159=0)
'      'Modify by Amy 2022/12/14 +e符號 Nvl(EState,'')
'      strSql = strSql + " union select ' ' AS V,'△'||decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','')||Nvl(EState,'') AS 本所案號 , DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'               "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort,C1.CU01||C1.CU127 CNT" & SeColPA & _
'               " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5, CUSTOMER C6,caseprogress,(" & strSQLE(1) & ") " & _
'               "WHERE pa09=na01(+) and " & IIf(strAppNo = strAppNo1, "CP55='" & strAppNo, "CP55>='" & strAppNo & "' and CP55<='" & strAppNo1) & "' " & strWhereCP & " And ((PA26<>'" & Me.Tag & "' Or PA26 Is Null) And (PA27<>'" & Me.Tag & "' Or PA27 Is Null) And (PA28<>'" & Me.Tag & "' Or PA28 Is Null) And (PA29<>'" & Me.Tag & "' Or PA29 Is Null) And (PA30<>'" & Me.Tag & "' Or PA30 Is Null)) and pa01 in (" & GetAddStr(Str02) & ") and pa04='00' and substr(CP55,1,8)=c1.cu01(+) and decode(substr(CP55,9,1),null,'0',substr(CP55,9,1))=c1.cu02(+) " & _
'               " AND SUBSTR(PA26,1,8)=C6.CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=C6.CU02(+) and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'               " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & strSqlEW(1) & strSQL1
'      strSql = strSql + " union select ' ' AS V,'△'||decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','')||Nvl(EState,'') AS 本所案號 , DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'               "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort,C1.CU01||C1.CU127 CNT" & SeColPA & _
'               " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5, CUSTOMER C6,caseprogress,(" & strSQLE(1) & ") " & _
'               "WHERE pa09=na01(+) and " & IIf(strAppNo = strAppNo1, "CP56='" & strAppNo, "CP56>='" & strAppNo & "' and CP56<='" & strAppNo1) & "' " & strWhereCP & " And ((PA26<>'" & Me.Tag & "' Or PA26 Is Null) And (PA27<>'" & Me.Tag & "' Or PA27 Is Null) And (PA28<>'" & Me.Tag & "' Or PA28 Is Null) And (PA29<>'" & Me.Tag & "' Or PA29 Is Null) And (PA30<>'" & Me.Tag & "' Or PA30 Is Null)) and pa01 in (" & GetAddStr(Str02) & ") and pa04='00' and substr(CP56,1,8)=c1.cu01(+) and decode(substr(CP56,9,1),null,'0',substr(CP56,9,1))=c1.cu02(+) " & _
'               " AND SUBSTR(PA26,1,8)=C6.CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=C6.CU02(+) and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'               " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & strSqlEW(1) & strSQL1
'      strSql = strSql + " union select ' ' AS V,'△'||decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','') ||Nvl(EState,'')AS 本所案號 , DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'               "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort,C1.CU01||C1.CU127 CNT" & SeColPA & _
'               " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5, CUSTOMER C6,caseprogress,(" & strSQLE(1) & ") " & _
'               "WHERE pa09=na01(+) and " & IIf(strAppNo = strAppNo1, "CP72='" & strAppNo, "CP72>='" & strAppNo & "' and CP72<='" & strAppNo1) & "' And ((PA26<>'" & Me.Tag & "' Or PA26 Is Null) And (PA27<>'" & Me.Tag & "' Or PA27 Is Null) And (PA28<>'" & Me.Tag & "' Or PA28 Is Null) And (PA29<>'" & Me.Tag & "' Or PA29 Is Null) And (PA30<>'" & Me.Tag & "' Or PA30 Is Null)) and pa01 in (" & GetAddStr(Str02) & ") and pa04='00' and substr(CP72,1,8)=c1.cu01(+) and decode(substr(CP72,9,1),null,'0',substr(CP72,9,1))=c1.cu02(+) " & _
'               " AND SUBSTR(PA26,1,8)=C6.CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=C6.CU02(+) and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'               " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & strSqlEW(1) & strSQL1
'      '2008/7/31 add by sonia 加案件進度檔 cp89~cp96
'      strSql = strSql + " union select ' ' AS V,'△'||decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','')||Nvl(EState,'') AS 本所案號 , DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'               "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort,C1.CU01||C1.CU127 CNT" & SeColPA & _
'               " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5, CUSTOMER C6,caseprogress,(" & strSQLE(1) & ") " & _
'               "WHERE pa09=na01(+) and " & IIf(strAppNo = strAppNo1, "CP89='" & strAppNo, "CP89>='" & strAppNo & "' and CP89<='" & strAppNo1) & "' " & strWhereCP & " And ((PA26<>'" & Me.Tag & "' Or PA26 Is Null) And (PA27<>'" & Me.Tag & "' Or PA27 Is Null) And (PA28<>'" & Me.Tag & "' Or PA28 Is Null) And (PA29<>'" & Me.Tag & "' Or PA29 Is Null) And (PA30<>'" & Me.Tag & "' Or PA30 Is Null)) and pa01 in (" & GetAddStr(Str02) & ") and pa04='00' and substr(CP89,1,8)=c1.cu01(+) and decode(substr(CP89,9,1),null,'0',substr(CP89,9,1))=c1.cu02(+) " & _
'               " AND SUBSTR(PA26,1,8)=C6.CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=C6.CU02(+) and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'               " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & strSqlEW(1) & strSQL1
'      strSql = strSql + " union select ' ' AS V,'△'||decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','')||Nvl(EState,'') AS 本所案號 , DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'               "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort,C1.CU01||C1.CU127 CNT" & SeColPA & _
'               " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5, CUSTOMER C6,caseprogress,(" & strSQLE(1) & ") " & _
'               "WHERE pa09=na01(+) and " & IIf(strAppNo = strAppNo1, "CP90='" & strAppNo, "CP90>='" & strAppNo & "' and CP90<='" & strAppNo1) & "' " & strWhereCP & " And ((PA26<>'" & Me.Tag & "' Or PA26 Is Null) And (PA27<>'" & Me.Tag & "' Or PA27 Is Null) And (PA28<>'" & Me.Tag & "' Or PA28 Is Null) And (PA29<>'" & Me.Tag & "' Or PA29 Is Null) And (PA30<>'" & Me.Tag & "' Or PA30 Is Null)) and pa01 in (" & GetAddStr(Str02) & ") and pa04='00' and substr(CP90,1,8)=c1.cu01(+) and decode(substr(CP90,9,1),null,'0',substr(CP90,9,1))=c1.cu02(+) " & _
'               " AND SUBSTR(PA26,1,8)=C6.CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=C6.CU02(+) and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'               " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & strSqlEW(1) & strSQL1
'      strSql = strSql + " union select ' ' AS V,'△'||decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','')||Nvl(EState,'') AS 本所案號 , DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'               "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort,C1.CU01||C1.CU127 CNT" & SeColPA & _
'               " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5, CUSTOMER C6,caseprogress,(" & strSQLE(1) & ") " & _
'               "WHERE pa09=na01(+) and " & IIf(strAppNo = strAppNo1, "CP91='" & strAppNo, "CP91>='" & strAppNo & "' and CP91<='" & strAppNo1) & "' " & strWhereCP & " And ((PA26<>'" & Me.Tag & "' Or PA26 Is Null) And (PA27<>'" & Me.Tag & "' Or PA27 Is Null) And (PA28<>'" & Me.Tag & "' Or PA28 Is Null) And (PA29<>'" & Me.Tag & "' Or PA29 Is Null) And (PA30<>'" & Me.Tag & "' Or PA30 Is Null)) and pa01 in (" & GetAddStr(Str02) & ") and pa04='00' and substr(CP91,1,8)=c1.cu01(+) and decode(substr(CP91,9,1),null,'0',substr(CP91,9,1))=c1.cu02(+) " & _
'               " AND SUBSTR(PA26,1,8)=C6.CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=C6.CU02(+) and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'               " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & strSqlEW(1) & strSQL1
'      strSql = strSql + " union select ' ' AS V,'△'||decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','')||Nvl(EState,'') AS 本所案號 , DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'               "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort,C1.CU01||C1.CU127 CNT" & SeColPA & _
'               " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5, CUSTOMER C6,caseprogress,(" & strSQLE(1) & ") " & _
'               "WHERE pa09=na01(+) and " & IIf(strAppNo = strAppNo1, "CP92='" & strAppNo, "CP92>='" & strAppNo & "' and CP92<='" & strAppNo1) & "' " & strWhereCP & " And ((PA26<>'" & Me.Tag & "' Or PA26 Is Null) And (PA27<>'" & Me.Tag & "' Or PA27 Is Null) And (PA28<>'" & Me.Tag & "' Or PA28 Is Null) And (PA29<>'" & Me.Tag & "' Or PA29 Is Null) And (PA30<>'" & Me.Tag & "' Or PA30 Is Null)) and pa01 in (" & GetAddStr(Str02) & ") and pa04='00' and substr(CP92,1,8)=c1.cu01(+) and decode(substr(CP92,9,1),null,'0',substr(CP92,9,1))=c1.cu02(+) " & _
'               " AND SUBSTR(PA26,1,8)=C6.CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=C6.CU02(+) and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'               " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & strSqlEW(1) & strSQL1
'      strSql = strSql + " union select ' ' AS V,'△'||decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','')||Nvl(EState,'') AS 本所案號 , DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'               "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort,C1.CU01||C1.CU127 CNT" & SeColPA & _
'               " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5, CUSTOMER C6,caseprogress,(" & strSQLE(1) & ") " & _
'               "WHERE pa09=na01(+) and " & IIf(strAppNo = strAppNo1, "CP93='" & strAppNo, "CP93>='" & strAppNo & "' and CP93<='" & strAppNo1) & "' " & strWhereCP & " And ((PA26<>'" & Me.Tag & "' Or PA26 Is Null) And (PA27<>'" & Me.Tag & "' Or PA27 Is Null) And (PA28<>'" & Me.Tag & "' Or PA28 Is Null) And (PA29<>'" & Me.Tag & "' Or PA29 Is Null) And (PA30<>'" & Me.Tag & "' Or PA30 Is Null)) and pa01 in (" & GetAddStr(Str02) & ") and pa04='00' and substr(CP93,1,8)=c1.cu01(+) and decode(substr(CP93,9,1),null,'0',substr(CP93,9,1))=c1.cu02(+) " & _
'               " AND SUBSTR(PA26,1,8)=C6.CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=C6.CU02(+) and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'               " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & strSqlEW(1) & strSQL1
'      strSql = strSql + " union select ' ' AS V,'△'||decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','')||Nvl(EState,'') AS 本所案號 , DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'               "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort,C1.CU01||C1.CU127 CNT" & SeColPA & _
'               " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5, CUSTOMER C6,caseprogress,(" & strSQLE(1) & ") " & _
'               "WHERE pa09=na01(+) and " & IIf(strAppNo = strAppNo1, "CP94='" & strAppNo, "CP94>='" & strAppNo & "' and CP94<='" & strAppNo1) & "' " & strWhereCP & " And ((PA26<>'" & Me.Tag & "' Or PA26 Is Null) And (PA27<>'" & Me.Tag & "' Or PA27 Is Null) And (PA28<>'" & Me.Tag & "' Or PA28 Is Null) And (PA29<>'" & Me.Tag & "' Or PA29 Is Null) And (PA30<>'" & Me.Tag & "' Or PA30 Is Null)) and pa01 in (" & GetAddStr(Str02) & ") and pa04='00' and substr(CP94,1,8)=c1.cu01(+) and decode(substr(CP94,9,1),null,'0',substr(CP94,9,1))=c1.cu02(+) " & _
'               " AND SUBSTR(PA26,1,8)=C6.CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=C6.CU02(+) and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'               " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & strSqlEW(1) & strSQL1
'      strSql = strSql + " union select ' ' AS V,'△'||decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','')||Nvl(EState,'') AS 本所案號 , DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'               "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort,C1.CU01||C1.CU127 CNT" & SeColPA & _
'               " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5, CUSTOMER C6,caseprogress,(" & strSQLE(1) & ") " & _
'               "WHERE pa09=na01(+) and " & IIf(strAppNo = strAppNo1, "CP95='" & strAppNo, "CP95>='" & strAppNo & "' and CP95<='" & strAppNo1) & "' " & strWhereCP & " And ((PA26<>'" & Me.Tag & "' Or PA26 Is Null) And (PA27<>'" & Me.Tag & "' Or PA27 Is Null) And (PA28<>'" & Me.Tag & "' Or PA28 Is Null) And (PA29<>'" & Me.Tag & "' Or PA29 Is Null) And (PA30<>'" & Me.Tag & "' Or PA30 Is Null)) and pa01 in (" & GetAddStr(Str02) & ") and pa04='00' and substr(CP95,1,8)=c1.cu01(+) and decode(substr(CP95,9,1),null,'0',substr(CP95,9,1))=c1.cu02(+) " & _
'               " AND SUBSTR(PA26,1,8)=C6.CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=C6.CU02(+) and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'               " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & strSqlEW(1) & strSQL1
'      strSql = strSql + " union select ' ' AS V,'△'||decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','')||Nvl(EState,'') AS 本所案號 , DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'               "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort,C1.CU01||C1.CU127 CNT" & SeColPA & _
'               " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5, CUSTOMER C6,caseprogress,(" & strSQLE(1) & ") " & _
'               "WHERE pa09=na01(+) and " & IIf(strAppNo = strAppNo1, "CP96='" & strAppNo, "CP96>='" & strAppNo & "' and CP96<='" & strAppNo1) & "' " & strWhereCP & " And ((PA26<>'" & Me.Tag & "' Or PA26 Is Null) And (PA27<>'" & Me.Tag & "' Or PA27 Is Null) And (PA28<>'" & Me.Tag & "' Or PA28 Is Null) And (PA29<>'" & Me.Tag & "' Or PA29 Is Null) And (PA30<>'" & Me.Tag & "' Or PA30 Is Null)) and pa01 in (" & GetAddStr(Str02) & ") and pa04='00' and substr(CP96,1,8)=c1.cu01(+) and decode(substr(CP96,9,1),null,'0',substr(CP96,9,1))=c1.cu02(+) " & _
'               " AND SUBSTR(PA26,1,8)=C6.CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=C6.CU02(+) and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'               " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & strSqlEW(1) & strSQL1
'
'      'Modified by Lydia 2019/12/26 增加欄位SeColSP
'      'Modify by Amy 2020/02/05 +SP73 商品類別
'      'Modify by Amy 2022/11/14 +strWhereCP 以X編號抓進度檔的CP55/CP56/CP89~CP96的資料時,要加入條件(CP158>0 OR CP159=0)
'      strSql = strSql + " union select ' ' AS V,'△'||SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'               "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort,C1.CU01||C1.CU127 CNT" & SeColSP & _
'               " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress WHERE sp09=na01(+) and " & IIf(strAppNo = strAppNo1, "CP55='" & strAppNo, "CP55>='" & strAppNo & "' and CP55<='" & strAppNo1) & "' " & strWhereCP & " And ((SP08<>'" & Me.Tag & "' Or SP08 Is Null ) And (SP58<>'" & Me.Tag & "' Or SP58 Is Null) And (SP59<>'" & Me.Tag & "' Or SP59 Is Null) And (SP65<>'" & Me.Tag & "' Or SP65 Is Null) And (SP66<>'" & Me.Tag & "' Or SP66 Is Null)) and sp01 in (" & GetAddStr(Str02) & ") AND SP04='00' AND SUBSTR(CP55,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP55,9,1),NULL,'0',SUBSTR(CP55,9,1))=C1.CU02(+) " & _
'               " AND SUBSTR(SP08,1,8)=C6.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C6.CU02(+) " & _
'               " AND SUBSTR(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) " & _
'               " AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'               " AND SUBSTR(SP65,1,8)=C4.CU01(+) AND DECODE(SUBSTR(SP65,9,1),NULL,'0',SUBSTR(SP65,9,1))=C4.CU02(+) " & _
'               " AND SUBSTR(SP66,1,8)=C5.CU01(+) AND DECODE(SUBSTR(SP66,9,1),NULL,'0',SUBSTR(SP66,9,1))=C5.CU02(+) and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & strSQL1
'      strSql = strSql + " union select ' ' AS V,'△'||SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'               "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort ,C1.CU01||C1.CU127 CNT" & SeColSP & _
'               " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress WHERE sp09=na01(+) and " & IIf(strAppNo = strAppNo1, "CP56='" & strAppNo, "CP56>='" & strAppNo & "' and CP56<='" & strAppNo1) & "' " & strWhereCP & " And ((SP08<>'" & Me.Tag & "' Or SP08 Is Null ) And (SP58<>'" & Me.Tag & "' Or SP58 Is Null) And (SP59<>'" & Me.Tag & "' Or SP59 Is Null) And (SP65<>'" & Me.Tag & "' Or SP65 Is Null) And (SP66<>'" & Me.Tag & "' Or SP66 Is Null)) and sp01 in (" & GetAddStr(Str02) & ") AND SP04='00' AND SUBSTR(CP56,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP56,9,1),NULL,'0',SUBSTR(CP56,9,1))=C1.CU02(+) " & _
'               " AND SUBSTR(SP08,1,8)=C6.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C6.CU02(+) " & _
'               " AND SUBSTR(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) " & _
'               " AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'               " AND SUBSTR(SP65,1,8)=C4.CU01(+) AND DECODE(SUBSTR(SP65,9,1),NULL,'0',SUBSTR(SP65,9,1))=C4.CU02(+) " & _
'               " AND SUBSTR(SP66,1,8)=C5.CU01(+) AND DECODE(SUBSTR(SP66,9,1),NULL,'0',SUBSTR(SP66,9,1))=C5.CU02(+) and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & strSQL1
'      strSql = strSql + " union select ' ' AS V,'△'||SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'               "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort ,C1.CU01||C1.CU127 CNT" & SeColSP & _
'               " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress WHERE sp09=na01(+) and " & IIf(strAppNo = strAppNo1, "CP72='" & strAppNo, "CP72>='" & strAppNo & "' and CP72<='" & strAppNo1) & "' And ((SP08<>'" & Me.Tag & "' Or SP08 Is Null ) And (SP58<>'" & Me.Tag & "' Or SP58 Is Null) And (SP59<>'" & Me.Tag & "' Or SP59 Is Null) And (SP65<>'" & Me.Tag & "' Or SP65 Is Null) And (SP66<>'" & Me.Tag & "' Or SP66 Is Null)) and sp01 in (" & GetAddStr(Str02) & ") AND SP04='00' AND SUBSTR(CP72,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP72,9,1),NULL,'0',SUBSTR(CP72,9,1))=C1.CU02(+) " & _
'               " AND SUBSTR(SP08,1,8)=C6.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C6.CU02(+) " & _
'               " AND SUBSTR(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) " & _
'               " AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'               " AND SUBSTR(SP65,1,8)=C4.CU01(+) AND DECODE(SUBSTR(SP65,9,1),NULL,'0',SUBSTR(SP65,9,1))=C4.CU02(+) " & _
'               " AND SUBSTR(SP66,1,8)=C5.CU01(+) AND DECODE(SUBSTR(SP66,9,1),NULL,'0',SUBSTR(SP66,9,1))=C5.CU02(+) and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & strSQL1
'      '2008/7/31 add by sonia 加案件進度檔 cp89~cp96
'      strSql = strSql + " union select ' ' AS V,'△'||SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'               "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort ,C1.CU01||C1.CU127 CNT" & SeColSP & _
'               " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress WHERE sp09=na01(+) and " & IIf(strAppNo = strAppNo1, "CP89='" & strAppNo, "CP89>='" & strAppNo & "' and CP89<='" & strAppNo1) & "' " & strWhereCP & " And ((SP08<>'" & Me.Tag & "' Or SP08 Is Null ) And (SP58<>'" & Me.Tag & "' Or SP58 Is Null) And (SP59<>'" & Me.Tag & "' Or SP59 Is Null) And (SP65<>'" & Me.Tag & "' Or SP65 Is Null) And (SP66<>'" & Me.Tag & "' Or SP66 Is Null)) and sp01 in (" & GetAddStr(Str02) & ") AND SP04='00' AND SUBSTR(CP89,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP89,9,1),NULL,'0',SUBSTR(CP89,9,1))=C1.CU02(+) " & _
'               " AND SUBSTR(SP08,1,8)=C6.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C6.CU02(+) " & _
'               " AND SUBSTR(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) " & _
'               " AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'               " AND SUBSTR(SP65,1,8)=C4.CU01(+) AND DECODE(SUBSTR(SP65,9,1),NULL,'0',SUBSTR(SP65,9,1))=C4.CU02(+) " & _
'               " AND SUBSTR(SP66,1,8)=C5.CU01(+) AND DECODE(SUBSTR(SP66,9,1),NULL,'0',SUBSTR(SP66,9,1))=C5.CU02(+) and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & strSQL1
'      strSql = strSql + " union select ' ' AS V,'△'||SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'               "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort,C1.CU01||C1.CU127 CNT" & SeColSP & _
'               " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress WHERE sp09=na01(+) and " & IIf(strAppNo = strAppNo1, "CP90='" & strAppNo, "CP90>='" & strAppNo & "' and CP90<='" & strAppNo1) & "' " & strWhereCP & " And ((SP08<>'" & Me.Tag & "' Or SP08 Is Null ) And (SP58<>'" & Me.Tag & "' Or SP58 Is Null) And (SP59<>'" & Me.Tag & "' Or SP59 Is Null) And (SP65<>'" & Me.Tag & "' Or SP65 Is Null) And (SP66<>'" & Me.Tag & "' Or SP66 Is Null)) and sp01 in (" & GetAddStr(Str02) & ") AND SP04='00' AND SUBSTR(CP90,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP90,9,1),NULL,'0',SUBSTR(CP90,9,1))=C1.CU02(+) " & _
'               " AND SUBSTR(SP08,1,8)=C6.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C6.CU02(+) " & _
'               " AND SUBSTR(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) " & _
'               " AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'               " AND SUBSTR(SP65,1,8)=C4.CU01(+) AND DECODE(SUBSTR(SP65,9,1),NULL,'0',SUBSTR(SP65,9,1))=C4.CU02(+) " & _
'               " AND SUBSTR(SP66,1,8)=C5.CU01(+) AND DECODE(SUBSTR(SP66,9,1),NULL,'0',SUBSTR(SP66,9,1))=C5.CU02(+) and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & strSQL1
'      strSql = strSql + " union select ' ' AS V,'△'||SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'               "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort,C1.CU01||C1.CU127 CNT" & SeColSP & _
'               " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress WHERE sp09=na01(+) and " & IIf(strAppNo = strAppNo1, "CP91='" & strAppNo, "CP91>='" & strAppNo & "' and CP91<='" & strAppNo1) & "' " & strWhereCP & " And ((SP08<>'" & Me.Tag & "' Or SP08 Is Null ) And (SP58<>'" & Me.Tag & "' Or SP58 Is Null) And (SP59<>'" & Me.Tag & "' Or SP59 Is Null) And (SP65<>'" & Me.Tag & "' Or SP65 Is Null) And (SP66<>'" & Me.Tag & "' Or SP66 Is Null)) and sp01 in (" & GetAddStr(Str02) & ") AND SP04='00' AND SUBSTR(CP91,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP91,9,1),NULL,'0',SUBSTR(CP91,9,1))=C1.CU02(+) " & _
'               " AND SUBSTR(SP08,1,8)=C6.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C6.CU02(+) " & _
'               " AND SUBSTR(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) " & _
'               " AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'               " AND SUBSTR(SP65,1,8)=C4.CU01(+) AND DECODE(SUBSTR(SP65,9,1),NULL,'0',SUBSTR(SP65,9,1))=C4.CU02(+) " & _
'               " AND SUBSTR(SP66,1,8)=C5.CU01(+) AND DECODE(SUBSTR(SP66,9,1),NULL,'0',SUBSTR(SP66,9,1))=C5.CU02(+) and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & strSQL1
'      strSql = strSql + " union select ' ' AS V,'△'||SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'               "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort ,C1.CU01||C1.CU127 CNT" & SeColSP & _
'               " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress WHERE sp09=na01(+) and " & IIf(strAppNo = strAppNo1, "CP92='" & strAppNo, "CP92>='" & strAppNo & "' and CP92<='" & strAppNo1) & "' " & strWhereCP & " And ((SP08<>'" & Me.Tag & "' Or SP08 Is Null ) And (SP58<>'" & Me.Tag & "' Or SP58 Is Null) And (SP59<>'" & Me.Tag & "' Or SP59 Is Null) And (SP65<>'" & Me.Tag & "' Or SP65 Is Null) And (SP66<>'" & Me.Tag & "' Or SP66 Is Null)) and sp01 in (" & GetAddStr(Str02) & ") AND SP04='00' AND SUBSTR(CP92,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP92,9,1),NULL,'0',SUBSTR(CP92,9,1))=C1.CU02(+) " & _
'               " AND SUBSTR(SP08,1,8)=C6.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C6.CU02(+) " & _
'               " AND SUBSTR(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) " & _
'               " AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'               " AND SUBSTR(SP65,1,8)=C4.CU01(+) AND DECODE(SUBSTR(SP65,9,1),NULL,'0',SUBSTR(SP65,9,1))=C4.CU02(+) " & _
'               " AND SUBSTR(SP66,1,8)=C5.CU01(+) AND DECODE(SUBSTR(SP66,9,1),NULL,'0',SUBSTR(SP66,9,1))=C5.CU02(+) and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & strSQL1
'      strSql = strSql + " union select ' ' AS V,'△'||SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'               "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort,C1.CU01||C1.CU127 CNT" & SeColSP & _
'               " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress WHERE sp09=na01(+) and " & IIf(strAppNo = strAppNo1, "CP93='" & strAppNo, "CP93>='" & strAppNo & "' and CP93<='" & strAppNo1) & "' " & strWhereCP & " And ((SP08<>'" & Me.Tag & "' Or SP08 Is Null ) And (SP58<>'" & Me.Tag & "' Or SP58 Is Null) And (SP59<>'" & Me.Tag & "' Or SP59 Is Null) And (SP65<>'" & Me.Tag & "' Or SP65 Is Null) And (SP66<>'" & Me.Tag & "' Or SP66 Is Null)) and sp01 in (" & GetAddStr(Str02) & ") AND SP04='00' AND SUBSTR(CP93,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP93,9,1),NULL,'0',SUBSTR(CP93,9,1))=C1.CU02(+) " & _
'               " AND SUBSTR(SP08,1,8)=C6.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C6.CU02(+) " & _
'               " AND SUBSTR(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) " & _
'               " AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'               " AND SUBSTR(SP65,1,8)=C4.CU01(+) AND DECODE(SUBSTR(SP65,9,1),NULL,'0',SUBSTR(SP65,9,1))=C4.CU02(+) " & _
'               " AND SUBSTR(SP66,1,8)=C5.CU01(+) AND DECODE(SUBSTR(SP66,9,1),NULL,'0',SUBSTR(SP66,9,1))=C5.CU02(+) and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & strSQL1
'      strSql = strSql + " union select ' ' AS V,'△'||SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'               "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort ,C1.CU01||C1.CU127 CNT" & SeColSP & _
'               " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress WHERE sp09=na01(+) and " & IIf(strAppNo = strAppNo1, "CP94='" & strAppNo, "CP94>='" & strAppNo & "' and CP94<='" & strAppNo1) & "' " & strWhereCP & " And ((SP08<>'" & Me.Tag & "' Or SP08 Is Null ) And (SP58<>'" & Me.Tag & "' Or SP58 Is Null) And (SP59<>'" & Me.Tag & "' Or SP59 Is Null) And (SP65<>'" & Me.Tag & "' Or SP65 Is Null) And (SP66<>'" & Me.Tag & "' Or SP66 Is Null)) and sp01 in (" & GetAddStr(Str02) & ") AND SP04='00' AND SUBSTR(CP94,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP94,9,1),NULL,'0',SUBSTR(CP94,9,1))=C1.CU02(+) " & _
'               " AND SUBSTR(SP08,1,8)=C6.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C6.CU02(+) " & _
'               " AND SUBSTR(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) " & _
'               " AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'               " AND SUBSTR(SP65,1,8)=C4.CU01(+) AND DECODE(SUBSTR(SP65,9,1),NULL,'0',SUBSTR(SP65,9,1))=C4.CU02(+) " & _
'               " AND SUBSTR(SP66,1,8)=C5.CU01(+) AND DECODE(SUBSTR(SP66,9,1),NULL,'0',SUBSTR(SP66,9,1))=C5.CU02(+) and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & strSQL1
'      strSql = strSql + " union select ' ' AS V,'△'||SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'               "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort,C1.CU01||C1.CU127 CNT" & SeColSP & _
'               " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress WHERE sp09=na01(+) and " & IIf(strAppNo = strAppNo1, "CP95='" & strAppNo, "CP95>='" & strAppNo & "' and CP95<='" & strAppNo1) & "' " & strWhereCP & " And ((SP08<>'" & Me.Tag & "' Or SP08 Is Null ) And (SP58<>'" & Me.Tag & "' Or SP58 Is Null) And (SP59<>'" & Me.Tag & "' Or SP59 Is Null) And (SP65<>'" & Me.Tag & "' Or SP65 Is Null) And (SP66<>'" & Me.Tag & "' Or SP66 Is Null)) and sp01 in (" & GetAddStr(Str02) & ") AND SP04='00' AND SUBSTR(CP95,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP95,9,1),NULL,'0',SUBSTR(CP95,9,1))=C1.CU02(+) " & _
'               " AND SUBSTR(SP08,1,8)=C6.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C6.CU02(+) " & _
'               " AND SUBSTR(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) " & _
'               " AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'               " AND SUBSTR(SP65,1,8)=C4.CU01(+) AND DECODE(SUBSTR(SP65,9,1),NULL,'0',SUBSTR(SP65,9,1))=C4.CU02(+) " & _
'               " AND SUBSTR(SP66,1,8)=C5.CU01(+) AND DECODE(SUBSTR(SP66,9,1),NULL,'0',SUBSTR(SP66,9,1))=C5.CU02(+) and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & strSQL1
'      strSql = strSql + " union select ' ' AS V,'△'||SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'               "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'               ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort,C1.CU01||C1.CU127 CNT" & SeColSP & _
'               " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress WHERE sp09=na01(+) and " & IIf(strAppNo = strAppNo1, "CP96='" & strAppNo, "CP96>='" & strAppNo & "' and CP96<='" & strAppNo1) & "' " & strWhereCP & " And ((SP08<>'" & Me.Tag & "' Or SP08 Is Null ) And (SP58<>'" & Me.Tag & "' Or SP58 Is Null) And (SP59<>'" & Me.Tag & "' Or SP59 Is Null) And (SP65<>'" & Me.Tag & "' Or SP65 Is Null) And (SP66<>'" & Me.Tag & "' Or SP66 Is Null)) and sp01 in (" & GetAddStr(Str02) & ") AND SP04='00' AND SUBSTR(CP96,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP96,9,1),NULL,'0',SUBSTR(CP96,9,1))=C1.CU02(+) " & _
'               " AND SUBSTR(SP08,1,8)=C6.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C6.CU02(+) " & _
'               " AND SUBSTR(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) " & _
'               " AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'               " AND SUBSTR(SP65,1,8)=C4.CU01(+) AND DECODE(SUBSTR(SP65,9,1),NULL,'0',SUBSTR(SP65,9,1))=C4.CU02(+) " & _
'               " AND SUBSTR(SP66,1,8)=C5.CU01(+) AND DECODE(SUBSTR(SP66,9,1),NULL,'0',SUBSTR(SP66,9,1))=C5.CU02(+) and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & strSQL1
'      'end 2020/02/05
'
'      'Modify By Sindy 2011/1/20 +LC43,LC44,LC45,LC46
'      'Modify by Amy 2018/09/17 加CasePropertyMap,不加專案服務案 lc52=Y 顯示 案件進度+案件性質
'      'Modified by Lydia 2019/12/26 增加欄位SeColLC
'      'Modify by Amy 2022/11/14 +strWhereCP 以X編號抓進度檔的CP55/CP56/CP89~CP96的資料時,要加入條件(CP158>0 OR CP159=0)
'      strSql = strSql + " union select ' ' AS V,'△'||LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort,C1.CU01||C1.CU127 CNT" & SeColLC & _
'               " FROM LAWCASE,nation,CUSTOMER C1, CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress,CasePropertyMap WHERE lc15=na01(+) and " & IIf(strAppNo = strAppNo1, "CP55='" & strAppNo, "CP55>='" & strAppNo & "' and CP55<='" & strAppNo1) & "' " & strWhereCP & " And (LC11<>'" & Me.Tag & "' Or LC11 Is Null) And (LC43<>'" & Me.Tag & "' Or LC43 Is Null) And (LC44<>'" & Me.Tag & "' Or LC44 Is Null) And (LC45<>'" & Me.Tag & "' Or LC45 Is Null) And (LC46<>'" & Me.Tag & "' Or LC46 Is Null) and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00'" & _
'               " AND SUBSTR(CP55,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP55,9,1),NULL,'0',SUBSTR(CP55,9,1))=C1.CU02(+) " & _
'               " AND SUBSTR(LC11,1,8)=C6.CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1))=C6.CU02(+) " & _
'               " AND SUBSTR(LC43,1,8)=C2.CU01(+) AND DECODE(SUBSTR(LC43,9,1),NULL,'0',SUBSTR(LC43,9,1))=C2.CU02(+) " & _
'               " AND SUBSTR(LC44,1,8)=C3.CU01(+) AND DECODE(SUBSTR(LC44,9,1),NULL,'0',SUBSTR(LC44,9,1))=C3.CU02(+) " & _
'               " AND SUBSTR(LC45,1,8)=C4.CU01(+) AND DECODE(SUBSTR(LC45,9,1),NULL,'0',SUBSTR(LC45,9,1))=C4.CU02(+) " & _
'               " AND SUBSTR(LC46,1,8)=C5.CU01(+) AND DECODE(SUBSTR(LC46,9,1),NULL,'0',SUBSTR(LC46,9,1))=C5.CU02(+) " & _
'               " and CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1
'      strSql = strSql + " union select ' ' AS V,'△'||LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort,C1.CU01||C1.CU127 CNT" & SeColLC & _
'               " FROM LAWCASE,nation,CUSTOMER C1, CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress,CasePropertyMap WHERE lc15=na01(+) and " & IIf(strAppNo = strAppNo1, "CP56='" & strAppNo, "CP56>='" & strAppNo & "' and CP56<='" & strAppNo1) & "' " & strWhereCP & " And (LC11<>'" & Me.Tag & "' Or LC11 Is Null) And (LC43<>'" & Me.Tag & "' Or LC43 Is Null) And (LC44<>'" & Me.Tag & "' Or LC44 Is Null) And (LC45<>'" & Me.Tag & "' Or LC45 Is Null) And (LC46<>'" & Me.Tag & "' Or LC46 Is Null) and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00'" & _
'               " AND SUBSTR(CP56,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP56,9,1),NULL,'0',SUBSTR(CP56,9,1))=C1.CU02(+) " & _
'               " AND SUBSTR(LC11,1,8)=C6.CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1))=C6.CU02(+) " & _
'               " AND SUBSTR(LC43,1,8)=C2.CU01(+) AND DECODE(SUBSTR(LC43,9,1),NULL,'0',SUBSTR(LC43,9,1))=C2.CU02(+) " & _
'               " AND SUBSTR(LC44,1,8)=C3.CU01(+) AND DECODE(SUBSTR(LC44,9,1),NULL,'0',SUBSTR(LC44,9,1))=C3.CU02(+) " & _
'               " AND SUBSTR(LC45,1,8)=C4.CU01(+) AND DECODE(SUBSTR(LC45,9,1),NULL,'0',SUBSTR(LC45,9,1))=C4.CU02(+) " & _
'               " AND SUBSTR(LC46,1,8)=C5.CU01(+) AND DECODE(SUBSTR(LC46,9,1),NULL,'0',SUBSTR(LC46,9,1))=C5.CU02(+) " & _
'               " and CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1
'      strSql = strSql + " union select ' ' AS V,'△'||LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort,C1.CU01||C1.CU127 CNT" & SeColLC & _
'               " FROM LAWCASE,nation,CUSTOMER C1, CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress,CasePropertyMap WHERE lc15=na01(+) and " & IIf(strAppNo = strAppNo1, "CP72='" & strAppNo, "CP72>='" & strAppNo & "' and CP72<='" & strAppNo1) & "' And (LC11<>'" & Me.Tag & "' Or LC11 Is Null) And (LC43<>'" & Me.Tag & "' Or LC43 Is Null) And (LC44<>'" & Me.Tag & "' Or LC44 Is Null) And (LC45<>'" & Me.Tag & "' Or LC45 Is Null) And (LC46<>'" & Me.Tag & "' Or LC46 Is Null) and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00'" & _
'               " AND SUBSTR(CP72,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP72,9,1),NULL,'0',SUBSTR(CP72,9,1))=C1.CU02(+) " & _
'               " AND SUBSTR(LC11,1,8)=C6.CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1))=C6.CU02(+) " & _
'               " AND SUBSTR(LC43,1,8)=C2.CU01(+) AND DECODE(SUBSTR(LC43,9,1),NULL,'0',SUBSTR(LC43,9,1))=C2.CU02(+) " & _
'               " AND SUBSTR(LC44,1,8)=C3.CU01(+) AND DECODE(SUBSTR(LC44,9,1),NULL,'0',SUBSTR(LC44,9,1))=C3.CU02(+) " & _
'               " AND SUBSTR(LC45,1,8)=C4.CU01(+) AND DECODE(SUBSTR(LC45,9,1),NULL,'0',SUBSTR(LC45,9,1))=C4.CU02(+) " & _
'               " AND SUBSTR(LC46,1,8)=C5.CU01(+) AND DECODE(SUBSTR(LC46,9,1),NULL,'0',SUBSTR(LC46,9,1))=C5.CU02(+) " & _
'               " and CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1
'
'      'Modify By Sindy 2011/1/20 +HC24,HC25,HC26,HC27
'      'Modified by Lydia 2019/12/26 增加欄位SeColHC
'      'Modify by Amy 2022/11/14 +strWhereCP 以X編號抓進度檔的CP55/CP56/CP89~CP96的資料時,要加入條件(CP158>0 OR CP159=0)
'      strSql = strSql + " union select ' ' AS V,'△'||HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,'台灣' AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort,C1.CU01||C1.CU127 CNT" & SeColHC & _
'               " FROM HIRECASE,CUSTOMER C1, CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress WHERE " & IIf(strAppNo = strAppNo1, "CP55='" & strAppNo, "CP55>='" & strAppNo & "' and CP55<='" & strAppNo1) & "' " & strWhereCP & " And (HC05<>'" & Me.Tag & "' Or HC05 Is Null) And (HC24<>'" & Me.Tag & "' Or HC24 Is Null) And (HC25<>'" & Me.Tag & "' Or HC25 Is Null) And (HC26<>'" & Me.Tag & "' Or HC26 Is Null) And (HC27<>'" & Me.Tag & "' Or HC27 Is Null) " & _
'               " and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00'" & _
'               " AND SUBSTR(CP55,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP55,9,1),NULL,'0',SUBSTR(CP55,9,1))=C1.CU02(+) " & _
'               " AND SUBSTR(HC05,1,8)=C6.CU01(+) AND DECODE(SUBSTR(HC05,9,1),NULL,'0',SUBSTR(HC05,9,1))=C6.CU02(+) " & _
'               " AND SUBSTR(HC24,1,8)=C2.CU01(+) AND DECODE(SUBSTR(HC24,9,1),NULL,'0',SUBSTR(HC24,9,1))=C2.CU02(+) " & _
'               " AND SUBSTR(HC25,1,8)=C3.CU01(+) AND DECODE(SUBSTR(HC25,9,1),NULL,'0',SUBSTR(HC25,9,1))=C3.CU02(+) " & _
'               " AND SUBSTR(HC26,1,8)=C4.CU01(+) AND DECODE(SUBSTR(HC26,9,1),NULL,'0',SUBSTR(HC26,9,1))=C4.CU02(+) " & _
'               " AND SUBSTR(HC27,1,8)=C5.CU01(+) AND DECODE(SUBSTR(HC27,9,1),NULL,'0',SUBSTR(HC27,9,1))=C5.CU02(+) " & _
'               " and CP01=HC01(+) and CP02=HC02(+) and CP03=HC03(+) and CP04=HC04(+) " & strSQL1
'      strSql = strSql + " union select ' ' AS V,'△'||HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,'台灣' AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort,C1.CU01||C1.CU127 CNT" & SeColHC & _
'               " FROM HIRECASE,CUSTOMER C1, CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress WHERE " & IIf(strAppNo = strAppNo1, "CP56='" & strAppNo, "CP56>='" & strAppNo & "' and CP56<='" & strAppNo1) & "' " & strWhereCP & " And (HC05<>'" & Me.Tag & "' Or HC05 Is Null) And (HC24<>'" & Me.Tag & "' Or HC24 Is Null) And (HC25<>'" & Me.Tag & "' Or HC25 Is Null) And (HC26<>'" & Me.Tag & "' Or HC26 Is Null) And (HC27<>'" & Me.Tag & "' Or HC27 Is Null) " & _
'               " and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00'" & _
'               " AND SUBSTR(CP56,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP56,9,1),NULL,'0',SUBSTR(CP56,9,1))=C1.CU02(+) " & _
'               " AND SUBSTR(HC05,1,8)=C6.CU01(+) AND DECODE(SUBSTR(HC05,9,1),NULL,'0',SUBSTR(HC05,9,1))=C6.CU02(+) " & _
'               " AND SUBSTR(HC24,1,8)=C2.CU01(+) AND DECODE(SUBSTR(HC24,9,1),NULL,'0',SUBSTR(HC24,9,1))=C2.CU02(+) " & _
'               " AND SUBSTR(HC25,1,8)=C3.CU01(+) AND DECODE(SUBSTR(HC25,9,1),NULL,'0',SUBSTR(HC25,9,1))=C3.CU02(+) " & _
'               " AND SUBSTR(HC26,1,8)=C4.CU01(+) AND DECODE(SUBSTR(HC26,9,1),NULL,'0',SUBSTR(HC26,9,1))=C4.CU02(+) " & _
'               " AND SUBSTR(HC27,1,8)=C5.CU01(+) AND DECODE(SUBSTR(HC27,9,1),NULL,'0',SUBSTR(HC27,9,1))=C5.CU02(+) " & _
'               " and CP01=HC01(+) and CP02=HC02(+) and CP03=HC03(+) and CP04=HC04(+) " & strSQL1
'      strSql = strSql + " union select ' ' AS V,'△'||HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,'台灣' AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort,C1.CU01||C1.CU127 CNT" & SeColHC & _
'               " FROM HIRECASE,CUSTOMER C1, CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,CUSTOMER C6,caseprogress WHERE " & IIf(strAppNo = strAppNo1, "CP72='" & strAppNo, "CP72>='" & strAppNo & "' and CP72<='" & strAppNo1) & "' And (HC05<>'" & Me.Tag & "' Or HC05 Is Null) And (HC24<>'" & Me.Tag & "' Or HC24 Is Null) And (HC25<>'" & Me.Tag & "' Or HC25 Is Null) And (HC26<>'" & Me.Tag & "' Or HC26 Is Null) And (HC27<>'" & Me.Tag & "' Or HC27 Is Null) " & _
'               " and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00'" & _
'               " AND SUBSTR(CP72,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP72,9,1),NULL,'0',SUBSTR(CP72,9,1))=C1.CU02(+) " & _
'               " AND SUBSTR(HC05,1,8)=C6.CU01(+) AND DECODE(SUBSTR(HC05,9,1),NULL,'0',SUBSTR(HC05,9,1))=C6.CU02(+) " & _
'               " AND SUBSTR(HC24,1,8)=C2.CU01(+) AND DECODE(SUBSTR(HC24,9,1),NULL,'0',SUBSTR(HC24,9,1))=C2.CU02(+) " & _
'               " AND SUBSTR(HC25,1,8)=C3.CU01(+) AND DECODE(SUBSTR(HC25,9,1),NULL,'0',SUBSTR(HC25,9,1))=C3.CU02(+) " & _
'               " AND SUBSTR(HC26,1,8)=C4.CU01(+) AND DECODE(SUBSTR(HC26,9,1),NULL,'0',SUBSTR(HC26,9,1))=C4.CU02(+) " & _
'               " AND SUBSTR(HC27,1,8)=C5.CU01(+) AND DECODE(SUBSTR(HC27,9,1),NULL,'0',SUBSTR(HC27,9,1))=C5.CU02(+) " & _
'               " and CP01=HC01(+) and CP02=HC02(+) and CP03=HC03(+) and CP04=HC04(+) " & strSQL1
'    'end 2022/11/14
'Else
'         '法務專用 add by nickc 2005/10/04
'         '2006/1/2 MODIFY BY SONIA 加總收文號欄,隱藏不顯示,點選顧問電話諮詢按鈕時傳參數時用
'         '2005/12/19 MODIFY BY SONIA 進度備註改案件名稱欄,法務案抓案件名稱,顧問案抓進度備註,取消結果欄,增回執日欄,依立卷問題3需求調整欄位位置
'         '2005/11/28 MODIFY BY SONIA 調整案件性質及相關人欄位位置
'         'Modified by Lydia 2015/10/05 '承辦律師'改為'承辦人'、'承辦法務'改為'協辦人員'
'         'Modify by Amy 2018/08/15 若專案服務案 lc52=Y 顯示 案件進度+案件性質
'         'Modified by Morgan 2019/1/30 SQLGrpStr(Str02, #)->GetAddStr(str02) 取消不必要系統別的檢查(減少語法執行次數,分所有較大影響)
'         'Modified by Lydia 2019/12/26 增加欄位SeColLC
'         strSql = "select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,Decode(LC52,'Y',CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')',LC05||' '||LC06||' '||LC07) AS 案件名稱 ,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人, sqldatet(cp05) AS 收文日,s2.st02 as 承辦人,s3.st02 as 協辦人員,sqldatet(cp27) as 發文日,sqldatet(cp46) as 回執日,s1.st02 as 智權人員,nvl(cpm03,cpm04) AS 案件性質,na03 AS 申請國家,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort,CP09,CU01||CU127 CNT" & SeColLC & _
'                  " FROM LAWCASE,nation,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,staff s3,fagent,SystemKind " & _
'                   " WHERE lc15=na01(+) and " & IIf(strAppNo = strAppNo1, "LC11='" & strAppNo, "LC11>='" & strAppNo & "' and LC11<='" & strAppNo1) & "' and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00' AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1)) = CU02(+) and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+)  and cp13=s1.st01(+) and cp14=s2.st01(+) and cp29=s3.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL1
'         'Add By Sindy 2011/1/20 +LC43,LC44,LC45,LC46
'         strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,Decode(LC52,'Y',CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')',LC05||' '||LC06||' '||LC07) AS 案件名稱 ,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人, sqldatet(cp05) AS 收文日,s2.st02 as 承辦人,s3.st02 as 協辦人員,sqldatet(cp27) as 發文日,sqldatet(cp46) as 回執日,s1.st02 as 智權人員,nvl(cpm03,cpm04) AS 案件性質,na03 AS 申請國家,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort,CP09,CU01||CU127 CNT" & SeColLC & _
'                  " FROM LAWCASE,nation,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,staff s3,fagent,SystemKind " & _
'                   " WHERE lc15=na01(+) and " & IIf(strAppNo = strAppNo1, "LC43='" & strAppNo, "LC43>='" & strAppNo & "' and LC43<='" & strAppNo1) & "' and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00' AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1)) = CU02(+) and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+)  and cp13=s1.st01(+) and cp14=s2.st01(+) and cp29=s3.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL1
'         strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,Decode(LC52,'Y',CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')',LC05||' '||LC06||' '||LC07) AS 案件名稱 ,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人, sqldatet(cp05) AS 收文日,s2.st02 as 承辦人,s3.st02 as 協辦人員,sqldatet(cp27) as 發文日,sqldatet(cp46) as 回執日,s1.st02 as 智權人員,nvl(cpm03,cpm04) AS 案件性質,na03 AS 申請國家,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort,CP09,CU01||CU127 CNT" & SeColLC & _
'                  " FROM LAWCASE,nation,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,staff s3,fagent,SystemKind " & _
'                   " WHERE lc15=na01(+) and " & IIf(strAppNo = strAppNo1, "LC44='" & strAppNo, "LC44>='" & strAppNo & "' and LC44<='" & strAppNo1) & "' and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00' AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1)) = CU02(+) and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+)  and cp13=s1.st01(+) and cp14=s2.st01(+) and cp29=s3.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL1
'         strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,Decode(LC52,'Y',CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')',LC05||' '||LC06||' '||LC07) AS 案件名稱 ,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人, sqldatet(cp05) AS 收文日,s2.st02 as 承辦人,s3.st02 as 協辦人員,sqldatet(cp27) as 發文日,sqldatet(cp46) as 回執日,s1.st02 as 智權人員,nvl(cpm03,cpm04) AS 案件性質,na03 AS 申請國家,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort,CP09,CU01||CU127 CNT" & SeColLC & _
'                  " FROM LAWCASE,nation,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,staff s3,fagent,SystemKind " & _
'                   " WHERE lc15=na01(+) and " & IIf(strAppNo = strAppNo1, "LC45='" & strAppNo, "LC45>='" & strAppNo & "' and LC45<='" & strAppNo1) & "' and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00' AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1)) = CU02(+) and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+)  and cp13=s1.st01(+) and cp14=s2.st01(+) and cp29=s3.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL1
'         strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,Decode(LC52,'Y',CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')',LC05||' '||LC06||' '||LC07) AS 案件名稱 ,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人, sqldatet(cp05) AS 收文日,s2.st02 as 承辦人,s3.st02 as 協辦人員,sqldatet(cp27) as 發文日,sqldatet(cp46) as 回執日,s1.st02 as 智權人員,nvl(cpm03,cpm04) AS 案件性質,na03 AS 申請國家,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort,CP09,CU01||CU127 CNT" & SeColLC & _
'                   " FROM LAWCASE,nation,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,staff s3,fagent,SystemKind " & _
'                   " WHERE lc15=na01(+) and " & IIf(strAppNo = strAppNo1, "LC46='" & strAppNo, "LC46>='" & strAppNo & "' and LC46<='" & strAppNo1) & "' and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00' AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1)) = CU02(+) and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+)  and cp13=s1.st01(+) and cp14=s2.st01(+) and cp29=s3.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL1
'
'         '2005/11/28 MODIFY BY SONIA 顧問期間改放在進度備註欄
'         'Modified by Lydia 2019/12/26 增加欄位SeColHC
'         strSql = strSql + " union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,decode(cp10,'0',(sqldatet(cp53)||'-'||sqldatet(cp54))||' ',CP64) AS 案件名稱 ,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人, sqldatet(cp05) AS 收文日,s2.st02 as 承辦人,s3.st02 as 協辦人員,sqldatet(cp27) as 發文日,sqldatet(cp46) as 回執日,s1.st02 as 智權人員,nvl(cpm03,cpm04) AS 案件性質,'台灣' AS 申請國家,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort,CP09,CU01||CU127 CNT" & SeColHC & _
'                  " FROM HIRECASE,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,staff s3,fagent,SystemKind " & _
'                   "  WHERE " & IIf(strAppNo = strAppNo1, "HC05='" & strAppNo, "HC05>='" & strAppNo & "' and HC05<='" & strAppNo1) & "' and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00' AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+) and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) and cp13=s1.st01(+) and cp14=s2.st01(+) and cp29=s3.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL1
'         'Add By Sindy 2011/1/20 +HC24,HC25,HC26,HC27
'         strSql = strSql + " union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,decode(cp10,'0',(sqldatet(cp53)||'-'||sqldatet(cp54))||' ',CP64) AS 案件名稱 ,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人, sqldatet(cp05) AS 收文日,s2.st02 as 承辦人,s3.st02 as 協辦人員,sqldatet(cp27) as 發文日,sqldatet(cp46) as 回執日,s1.st02 as 智權人員,nvl(cpm03,cpm04) AS 案件性質,'台灣' AS 申請國家,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort,CP09,CU01||CU127 CNT" & SeColHC & _
'                  " FROM HIRECASE,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,staff s3,fagent,SystemKind " & _
'                   "  WHERE " & IIf(strAppNo = strAppNo1, "HC24='" & strAppNo, "HC24>='" & strAppNo & "' and HC24<='" & strAppNo1) & "' and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00' AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+) and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) and cp13=s1.st01(+) and cp14=s2.st01(+) and cp29=s3.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL1
'         strSql = strSql + " union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,decode(cp10,'0',(sqldatet(cp53)||'-'||sqldatet(cp54))||' ',CP64) AS 案件名稱 ,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人, sqldatet(cp05) AS 收文日,s2.st02 as 承辦人,s3.st02 as 協辦人員,sqldatet(cp27) as 發文日,sqldatet(cp46) as 回執日,s1.st02 as 智權人員,nvl(cpm03,cpm04) AS 案件性質,'台灣' AS 申請國家,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort,CP09,CU01||CU127 CNT" & SeColHC & _
'                  " FROM HIRECASE,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,staff s3,fagent,SystemKind " & _
'                   "  WHERE " & IIf(strAppNo = strAppNo1, "HC25='" & strAppNo, "HC25>='" & strAppNo & "' and HC25<='" & strAppNo1) & "' and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00' AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+) and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) and cp13=s1.st01(+) and cp14=s2.st01(+) and cp29=s3.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL1
'         strSql = strSql + " union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,decode(cp10,'0',(sqldatet(cp53)||'-'||sqldatet(cp54))||' ',CP64) AS 案件名稱 ,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人, sqldatet(cp05) AS 收文日,s2.st02 as 承辦人,s3.st02 as 協辦人員,sqldatet(cp27) as 發文日,sqldatet(cp46) as 回執日,s1.st02 as 智權人員,nvl(cpm03,cpm04) AS 案件性質,'台灣' AS 申請國家,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort,CP09,CU01||CU127 CNT" & SeColHC & _
'                  " FROM HIRECASE,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,staff s3,fagent,SystemKind " & _
'                   "  WHERE " & IIf(strAppNo = strAppNo1, "HC26='" & strAppNo, "HC26>='" & strAppNo & "' and HC26<='" & strAppNo1) & "' and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00' AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+) and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) and cp13=s1.st01(+) and cp14=s2.st01(+) and cp29=s3.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL1
'         strSql = strSql + " union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,decode(cp10,'0',(sqldatet(cp53)||'-'||sqldatet(cp54))||' ',CP64) AS 案件名稱 ,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人, sqldatet(cp05) AS 收文日,s2.st02 as 承辦人,s3.st02 as 協辦人員,sqldatet(cp27) as 發文日,sqldatet(cp46) as 回執日,s1.st02 as 智權人員,nvl(cpm03,cpm04) AS 案件性質,'台灣' AS 申請國家,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort,CP09,CU01||CU127 CNT" & SeColHC & _
'                  " FROM HIRECASE,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,staff s3,fagent,SystemKind " & _
'                   "  WHERE " & IIf(strAppNo = strAppNo1, "HC27='" & strAppNo, "HC27>='" & strAppNo & "' and HC27<='" & strAppNo1) & "' and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00' AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+) and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) and cp13=s1.st01(+) and cp14=s2.st01(+) and cp29=s3.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL1
'
'         'Modify By Sindy 2011/1/20 +LC43,LC44,LC45,LC46
'         'Modify by Amy 2018/08/15 若專案服務案 lc52=Y 顯示 案件進度+案件性質
'         'Modified by Lydia 2019/12/26 增加欄位SeColLC
'         'Modify by Amy 2022/11/14 +strWhereCP 以X編號抓進度檔的CP55/CP56/CP89~CP96的資料時,要加入條件(CP158>0 OR CP159=0)
'         strSql = strSql + " union select ' ' AS V,'△'||LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,Decode(LC52,'Y',CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')',LC05||' '||LC06||' '||LC07) AS 案件名稱 ,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人, sqldatet(cp05) AS 收文日,s2.st02 as 承辦人,s3.st02 as 協辦人員,sqldatet(cp27) as 發文日,sqldatet(cp46) as 回執日,s1.st02 as 智權人員,nvl(cpm03,cpm04) AS 案件性質,na03 AS 申請國家,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人," & _
'                  " LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort,CP09,CU01||CU127 CNT" & SeColLC & _
'                  " FROM LAWCASE,nation,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,staff s3,fagent,SystemKind WHERE lc15=na01(+) and " & IIf(strAppNo = strAppNo1, "CP55='" & strAppNo, "CP55>='" & strAppNo & "' and CP55<='" & strAppNo1) & "' " & strWhereCP & " And (LC11<>'" & Me.Tag & "' Or LC11 Is Null) And (LC43<>'" & Me.Tag & "' Or LC43 Is Null) And (LC44<>'" & Me.Tag & "' Or LC44 Is Null) And (LC45<>'" & Me.Tag & "' Or LC45 Is Null) And (LC46<>'" & Me.Tag & "' Or LC46 Is Null) and cp01 in (" & GetAddStr(Str02) & ") AND cp04='00' AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+)  and CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) and lc01 is not null and cp13=s1.st01(+) and cp14=s2.st01(+) and cp29=s3.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL1
'         strSql = strSql + " union select ' ' AS V,'△'||LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,Decode(LC52,'Y',CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')',LC05||' '||LC06||' '||LC07) AS 案件名稱 ,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人, sqldatet(cp05) AS 收文日,s2.st02 as 承辦人,s3.st02 as 協辦人員,sqldatet(cp27) as 發文日,sqldatet(cp46) as 回執日,s1.st02 as 智權人員,nvl(cpm03,cpm04) AS 案件性質,na03 AS 申請國家,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人," & _
'                  " LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort,CP09,CU01||CU127 CNT" & SeColLC & _
'                  " FROM LAWCASE,nation,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,staff s3,fagent,SystemKind WHERE lc15=na01(+) and " & IIf(strAppNo = strAppNo1, "CP56='" & strAppNo, "CP56>='" & strAppNo & "' and CP56<='" & strAppNo1) & "' " & strWhereCP & " And (LC11<>'" & Me.Tag & "' Or LC11 Is Null) And (LC43<>'" & Me.Tag & "' Or LC43 Is Null) And (LC44<>'" & Me.Tag & "' Or LC44 Is Null) And (LC45<>'" & Me.Tag & "' Or LC45 Is Null) And (LC46<>'" & Me.Tag & "' Or LC46 Is Null) and cp01 in (" & GetAddStr(Str02) & ") AND cp04='00' AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+)  and CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) and lc01 is not null and cp13=s1.st01(+) and cp14=s2.st01(+) and cp29=s3.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL1
'         strSql = strSql + " union select ' ' AS V,'△'||LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,Decode(LC52,'Y',CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')',LC05||' '||LC06||' '||LC07) AS 案件名稱 ,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人, sqldatet(cp05) AS 收文日,s2.st02 as 承辦人,s3.st02 as 協辦人員,sqldatet(cp27) as 發文日,sqldatet(cp46) as 回執日,s1.st02 as 智權人員,nvl(cpm03,cpm04) AS 案件性質,na03 AS 申請國家,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人," & _
'                  " LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort,CP09,CU01||CU127 CNT" & SeColLC & _
'                  " FROM LAWCASE,nation,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,staff s3,fagent,SystemKind WHERE lc15=na01(+) and " & IIf(strAppNo = strAppNo1, "CP72='" & strAppNo, "CP72>='" & strAppNo & "' and CP72<='" & strAppNo1) & "' And (LC11<>'" & Me.Tag & "' Or LC11 Is Null) And (LC43<>'" & Me.Tag & "' Or LC43 Is Null) And (LC44<>'" & Me.Tag & "' Or LC44 Is Null) And (LC45<>'" & Me.Tag & "' Or LC45 Is Null) And (LC46<>'" & Me.Tag & "' Or LC46 Is Null) and cp01 in (" & GetAddStr(Str02) & ") AND cp04='00' AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+)  and CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) and lc01 is not null and cp13=s1.st01(+) and cp14=s2.st01(+) and cp29=s3.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL1
'
'         '2005/11/28 MODIFY BY SONIA 顧問期間改放在進度備註欄
'         'Modify By Sindy 2011/1/20 +HC24,HC25,HC26,HC27
'         'Modified by Lydia 2019/12/26 增加欄位SeColHC
'         'Modify by Amy 2022/11/14 +strWhereCP 以X編號抓進度檔的CP55/CP56/CP89~CP96的資料時,要加入條件(CP158>0 OR CP159=0)
'         strSql = strSql + " union select ' ' AS V,'△'||HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,decode(cp10,'0',(sqldatet(cp53)||'-'||sqldatet(cp54))||' ',CP64) AS 案件名稱 ,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人, sqldatet(cp05) AS 收文日,s2.st02 as 承辦人,s3.st02 as 協辦人員,sqldatet(cp27) as 發文日,sqldatet(cp46) as 回執日,s1.st02 as 智權人員,nvl(cpm03,cpm04) AS 案件性質,'台灣' AS 申請國家,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人," & _
'                  " HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort,CP09,CU01||CU127 CNT" & SeColHC & _
'                  " FROM HIRECASE,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,staff s3,fagent,SystemKind WHERE " & IIf(strAppNo = strAppNo1, "CP55='" & strAppNo, "CP55>='" & strAppNo & "' and CP55<='" & strAppNo1) & "' " & strWhereCP & " And (HC05<>'" & Me.Tag & "' Or HC05 Is Null) And (HC24<>'" & Me.Tag & "' Or HC24 Is Null) And (HC25<>'" & Me.Tag & "' Or HC25 Is Null) And (HC26<>'" & Me.Tag & "' Or HC26 Is Null) And (HC27<>'" & Me.Tag & "' Or HC27 Is Null) and cp01 in (" & GetAddStr(Str02) & ") AND cp04='00' AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+)  and CP01=HC01(+) and CP02=HC02(+) and CP03=HC03(+) and CP04=HC04(+) and hc01 is not null and cp13=s1.st01(+) and cp14=s2.st01(+) and cp29=s3.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL1
'         strSql = strSql + " union select ' ' AS V,'△'||HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,decode(cp10,'0',(sqldatet(cp53)||'-'||sqldatet(cp54))||' ',CP64) AS 案件名稱 ,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人, sqldatet(cp05) AS 收文日,s2.st02 as 承辦人,s3.st02 as 協辦人員,sqldatet(cp27) as 發文日,sqldatet(cp46) as 回執日,s1.st02 as 智權人員,nvl(cpm03,cpm04) AS 案件性質,'台灣' AS 申請國家,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人," & _
'                  " HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort,CP09,CU01||CU127 CNT" & SeColHC & _
'                  " FROM HIRECASE,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,staff s3,fagent,SystemKind WHERE " & IIf(strAppNo = strAppNo1, "CP56='" & strAppNo, "CP56>='" & strAppNo & "' and CP56<='" & strAppNo1) & "' " & strWhereCP & " And (HC05<>'" & Me.Tag & "' Or HC05 Is Null) And (HC24<>'" & Me.Tag & "' Or HC24 Is Null) And (HC25<>'" & Me.Tag & "' Or HC25 Is Null) And (HC26<>'" & Me.Tag & "' Or HC26 Is Null) And (HC27<>'" & Me.Tag & "' Or HC27 Is Null) and cp01 in (" & GetAddStr(Str02) & ") AND cp04='00' AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+)  and CP01=HC01(+) and CP02=HC02(+) and CP03=HC03(+) and CP04=HC04(+) and hc01 is not null and cp13=s1.st01(+) and cp14=s2.st01(+) and cp29=s3.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL1
'         strSql = strSql + " union select ' ' AS V,'△'||HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,decode(cp10,'0',(sqldatet(cp53)||'-'||sqldatet(cp54))||' ',CP64) AS 案件名稱 ,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人, sqldatet(cp05) AS 收文日,s2.st02 as 承辦人,s3.st02 as 協辦人員,sqldatet(cp27) as 發文日,sqldatet(cp46) as 回執日,s1.st02 as 智權人員,nvl(cpm03,cpm04) AS 案件性質,'台灣' AS 申請國家,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人," & _
'                  " HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort,CP09,CU01||CU127 CNT" & SeColHC & _
'                  " FROM HIRECASE,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,staff s3,fagent,SystemKind WHERE " & IIf(strAppNo = strAppNo1, "CP72='" & strAppNo, "CP72>='" & strAppNo & "' And CP72<='" & strAppNo1) & "' And (HC05<>'" & Me.Tag & "' Or HC05 Is Null) And (HC24<>'" & Me.Tag & "' Or HC24 Is Null) And (HC25<>'" & Me.Tag & "' Or HC25 Is Null) And (HC26<>'" & Me.Tag & "' Or HC26 Is Null) And (HC27<>'" & Me.Tag & "' Or HC27 Is Null) and cp01 in (" & GetAddStr(Str02) & ") AND cp04='00' AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+)  and CP01=HC01(+) and CP02=HC02(+) and CP03=HC03(+) and CP04=HC04(+) and hc01 is not null and cp13=s1.st01(+) and cp14=s2.st01(+) and cp29=s3.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL1
'End If
''add by nickc 2007/03/23 更換 PCT 欄
''edit b nickc 2007/12/21
''If frm100102_1.ChkPCT.Value = vbChecked Then
'If chkpct.Value = vbChecked Then
'    strSql = Replace(Replace(Replace(Replace(UCase(strSql), "TM15 AS 審定專利號數", "'' as PCT"), "NVL(SP14,SP13) AS 審定專利號數", "'' as PCT"), "' ' AS 審定專利號數", "'' as PCT"), "PA22 AS 審定專利號數", "pa46 as PCT")
'End If
''加國家範圍條件
'If m_Cty1 <> "" Or m_Cty2 <> "" Then
'   strSql = "Select X.* From (" & strSql & ") X,Nation Y Where na03(+)=申請國家"
'   If m_Cty1 <> "" Then
'      strSql = strSql & " and na01>='" & m_Cty1 & "'"
'   End If
'   If m_Cty2 <> "" Then
'      strSql = strSql & " and na01<='" & m_Cty2 & "'"
'   End If
'   If strContactNo <> "" Then
'      strSql = strSql & " AND CNT='" & Left(strAppNo, 8) & strContactNo & "'"
'   End If
'ElseIf strContactNo <> "" Then
'   strSql = "Select X.* From (" & strSql & ") X Where CNT='" & Left(strAppNo, 8) & strContactNo & "'"
'End If
''Add end
'
''edit by nickc  2005/05/06 改排序
''strSQL = strSQL & " ORDER BY 本所案號"
''2005/12/21 MODIFY BY SONIA 法務進度再加收文日
''strSQL = strSQL & " ORDER BY FSort,本所案號"
'If bolIsL = False Then
'   strSql = strSql & " ORDER BY FSort,本所案號"
'Else
'   strSql = strSql & " ORDER BY FSort,本所案號,收文日"
'End If
''2005/12/21 END
''Modify end 2003/12/19
'
''Added by Lydia 2019/11/01 利益衝突案件：處理替換字串
''Mark by Lydia 2019/12/26
''If m_CuFaArea <> "" And stConPA & stConSP <> "" Then
''    stCuFaSQL = strSql
''    stCuFaSQL = Replace(stCuFaSQL, "CUFA_PA", stConPA)
''    stCuFaSQL = Replace(stCuFaSQL, "CUFA_SP", stConSP)
''    intI = 1
''    Set rsCnt = Nothing
''    Set rsCnt = ClsLawReadRstMsg(intI, stCuFaSQL)
''End If
''strSql = Replace(strSql, "CUFA_PA", "")
''strSql = Replace(strSql, "CUFA_SP", "")
'''end 2019/11/01
''end 2019/12/26
'
'CheckOC
'adoRecordset.CursorLocation = adUseClient
''Modified by Lydia 2019/12/26 改變型態
''adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
'
'If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'    If Len(Trim(Str02)) <> 0 Then
'        strTemp = Split(Str02, ",")
'    End If
'
'    'Modified by Lydia 2019/12/26 利益衝突案件：逐案號判斷
'    'adoRecordset.MoveFirst
'    'Dim StrTest2 As String, StrTest4 As String, s As Integer
'    'Set m_adoRst = adoRecordset.Clone 'Added by Lydia 2018/02/09 'move by Lydia 2018/12/17 從下面移上來
'    'If adoRecordset.RecordCount = 0 Then
'    '    Me.Enabled = True
'    '    cmdOK(0).Enabled = False
'    '    cmdOK(1).Enabled = False
'    '    cmdOK(2).Enabled = False
'    '    cmdOK(3).Enabled = False
'    '    ShowNoData
'    '    Screen.MousePointer = vbDefault
'    '    Me.Enabled = True
'    '    '920416 nick
'    '    'Me.Hide
'    '    tmpBol = fnCancelNowFormAndShowParentForm(Me)
'    '    Exit Sub
'     If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
'        intCufaCnt = 0
'        adoRecordset.MoveFirst
'        Do While adoRecordset.EOF = False
'            '利益衝突案件：逐案號判斷
'            If PUB_ChkCufaByCase(Me.Name, Str02, "" & adoRecordset.Fields("本所案號"), "" & adoRecordset.Fields("cust01") & "," & adoRecordset.Fields("cust02") & "," & adoRecordset.Fields("cust03") & "," & adoRecordset.Fields("cust04") & "," & adoRecordset.Fields("cust05"), "" & adoRecordset.Fields("fcno")) = False Then
'                intCufaCnt = intCufaCnt + 1
'                adoRecordset.Delete
'            End If
'            adoRecordset.MoveNext
'        Loop
'        '利益衝突案件：限閱案件
'        If intCufaCnt > 0 Then
'            MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
'        End If
'        If adoRecordset.RecordCount = 0 Then
'              GoTo JumpToNoData
'        End If
'     End If
'     Set m_adoRst = adoRecordset.Clone
'    'end 2019/12/26
'Else
'JumpToNoData: 'Added by Lydia 2019/12/26
'    Set m_adoRst = adoRecordset.Clone 'Added by Lydia 2018/02/09
'    'Added by Lydia 2019/12/26
'    cmdOK(0).Enabled = False
'    cmdOK(1).Enabled = False
'    cmdOK(2).Enabled = False
'    cmdOK(3).Enabled = False
'    'end 2019/12/26
'    ShowNoData
'    Screen.MousePointer = vbDefault
'    Me.Enabled = True
'    '920416 nick
'    'Me.Hide
'    tmpBol = fnCancelNowFormAndShowParentForm(Me)
'    Exit Sub
'End If
'''Add by Morgan 2004/1/5
'GrdDataList.FixedCols = 0
''Modified by Lydia 2018/12/17 中所反應耗時過久,卡在丟暫存檔(O8的寫法); O12 可以直接排序
''Set GrdDataList.Recordset = adoRecordset
'''Added by Lydia 2018/02/09 放到暫存檔,供Grid排序
''Set m_adoRst = PUB_CreateRecordset(adoRecordset, , , 300, Me.Name)
''Modified by Lydia 2018/12/22 拿掉desc
''m_adoRst.Sort = "FSort desc,本所案號 asc" 'Move by Lydia 2018/12/17 先排序,後指定資料集
'm_adoRst.Sort = "FSort ,本所案號 asc"
'SetRst2Grid
''end 2018/12/17
'm_blnColOrderAsc = True
''end 2018/02/09
'
''add by nick 2004/07/07
'SetDataListWidth
''Add by Morgan 2004/1/5
''2005/12/19 MODIFY BY SONIA
''GrdDataList.FixedCols = 4
'If bolIsL = False Then
'   GrdDataList.FixedCols = 4
'Else
'   GrdDataList.FixedCols = 6
'End If
''2005/12/19 END
'Me.Enabled = True
'
''2006/1/2 ADD BY SONIA 有權限者顯示顧問電話諮詢按鈕
'cmdOK(8).Enabled = False
'cmdOK(8).Visible = False
'If bolIsL = True Then
'   cmdOK(8).Enabled = IsUserHasRightOfFunction("frm100102_1", strAdd, False)
'   cmdOK(8).Visible = IsUserHasRightOfFunction("frm100102_1", strAdd, False)
'End If
''2006/1/2 END
End Sub 'End StrMenu_Old 語法改共用函數

'由代理人來    nick 91.08.01
'Mark by Amy 2022/12/14 沒在用不維護
''Memo by Lydia 2019/11/01 (2008/11/27)已改用StrMenu
'Sub StrMenu2()
'BolFrom100114 = True
'Dim strSQL2 As String
'Dim StrSQL3 As String
'Dim StrSQL4 As String
'Dim strSQL5 As String
'Dim StrSQL6 As String
'Dim strSQL8 As String
'Me.Enabled = False
'Str01 = ""    '申請人編號
'Str02 = ""    '系統類別
'Str01 = Me.Tag
'
''Add By Sindy 2011/01/03 檢查國內外權限
'If CheckSR12(Str01) = False Then
'   Screen.MousePointer = vbDefault
'   Me.Enabled = True
'   tmpBol = fnCancelNowFormAndShowParentForm(Me)
'   Exit Sub
'End If
'
''組字串
'strSQL1 = ""
'strSQL2 = ""
'StrSQL3 = ""
'StrSQL4 = ""
'strSQL5 = ""
'StrSQL6 = ""
'If Len(Trim(m_Sys)) <> 0 Then
'   strSQL1 = strSQL1 & " and tm01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 2) & ") "
'   strSQL2 = strSQL2 & " and pa01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 1) & ") "
'   StrSQL3 = StrSQL3 & " and sp01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 5) & ") "
'   StrSQL4 = StrSQL4 & " and lc01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 3) & ") "
'   strSQL8 = strSQL8 & " and hc01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 4) & ") "
'End If
'If Len(Trim(m_Cty1)) <> 0 Then           '檢查申請國家
'   strSQL1 = strSQL1 + " AND TM10='" & m_Cty1 & "' "
'   strSQL2 = strSQL2 + " AND PA09='" & m_Cty1 & "' "
'   StrSQL3 = StrSQL3 + " AND SP09='" & m_Cty1 & "' "
'   StrSQL4 = StrSQL4 + " AND LC15='" & m_Cty1 & "' "
'End If
'If Len(Trim(m_Pty1)) <> 0 Then            '檢查案件性質
'strSQL1 = strSQL1 + " AND CP10>='" & m_Pty1 & "' "
'strSQL2 = strSQL2 + " AND CP10>='" & m_Pty1 & "' "
'StrSQL3 = StrSQL3 + " AND CP10>='" & m_Pty1 & "' "
'StrSQL4 = StrSQL4 + " AND CP10>='" & m_Pty1 & "' "
'strSQL8 = strSQL8 + " AND CP10>='" & m_Pty1 & "' "
'End If
'If Len(Trim(m_Pty2)) <> 0 Then
'strSQL1 = strSQL1 + " AND CP10<='" & m_Pty2 & "' "
'strSQL2 = strSQL2 + " AND CP10<='" & m_Pty2 & "' "
'StrSQL3 = StrSQL3 + " AND CP10<='" & m_Pty2 & "' "
'StrSQL4 = StrSQL4 + " AND CP10<='" & m_Pty2 & "' "
'strSQL8 = strSQL8 + " AND CP10<='" & m_Pty2 & "' "
'End If
'If m_Type = "1" Then        '收文
'   If Len(m_Date1) <> 0 Then
'      strSQL5 = strSQL5 + " AND CP05>=" & Val(ChangeTStringToWString(m_Date1)) & " "
'   End If
'   If Len(m_Date2) <> 0 Then
'      strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(m_Date2)) & " "
'   'Add By Cheng 2002/03/18
'   Else
'      If Len(m_Date1) > 0 Then
'         strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(ServerDate - 19110000)) & " "
'      End If
'   End If
'Else
'   If Len(m_Date1) <> 0 Then
'      strSQL5 = strSQL5 + " AND CP27>=" & Val(ChangeTStringToWString(m_Date1)) & " "
'   End If
'   If Len(m_Date2) <> 0 Then
'      strSQL5 = strSQL5 + " AND CP27<=" & Val(ChangeTStringToWString(m_Date2)) & " "
'   'Add By Cheng 2002/03/18
'   Else
'      If Len(m_Date1) > 0 Then
'         strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(ServerDate - 19110000)) & " "
'      End If
'   End If
'End If
'
''Added by Lydia 2019/11/01 非法務案+屬於利益衝突案件之XY編號
''Mark by Lydia 2019/12/26
''stConPA = "": stConSP = ""
''If bolIsL = False And strSrvDate(1) >= XY特殊權限啟用日 And InStr(XY特殊權限範圍, Left(ChangeCustomerL(Me.Tag), 8)) > 0 Then
''    cnnConnection.Execute "delete from R100102_2 where R02201='" & strUserNum & "' and R02202='" & Me.Name & "' " '清空暫存檔
''    If PUB_ChkCuFa_Right(Me.Name, Me.Tag, Str02, m_CuFaRight, m_CuFaArea) = True Then
''    End If
''    '有管制系統別=>組合SQL條件
''    If m_CuFaArea <> "" Then
''        stConPA = Pub_CufaConSQL(Me.Name, "PA", Me.Tag, m_CuFaRight, m_CuFaArea)
''        stConSP = Pub_CufaConSQL(Me.Name, "SP", Me.Tag, m_CuFaRight, m_CuFaArea)
''    End If
''End If
'''end 2019/11/01
''end 2019/12/26
'
''顯示表單上面的值
'Label3.Caption = Me.Tag
'If Len(Trim(Me.Tag)) = 9 Then
'   strSql = "SELECT NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),CU13,ST02,cu111 FROM CUSTOMER,STAFF WHERE CU01='" & Left$(GetNewFagent(Me.Tag), 8) & "' AND CU02='" & Right$(GetNewFagent(Me.Tag), 1) & "' AND CU13=ST01(+)"
'Else
'   strSql = "SELECT NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),CU13,ST02,cu111 FROM CUSTOMER,STAFF WHERE CU01='" & Left$(GetNewFagent(Me.Tag), 8) & "' AND CU02='0' AND CU13=ST01(+) "
'End If
'CheckOC
'adoRecordset.CursorLocation = adUseClient
'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'    If IsNull(adoRecordset.Fields(0)) Then
'        Label4.Caption = ""
'    Else
'        Label4.Caption = adoRecordset.Fields(0)
'    End If
'    If IsNull(adoRecordset.Fields(1)) Then
'        Label6.Caption = ""
'    Else
'        Label6.Caption = adoRecordset.Fields(1)
'    End If
'    If IsNull(adoRecordset.Fields(2)) Then
'        Label7.Caption = ""
'    Else
'        Label7.Caption = adoRecordset.Fields(2)
'    End If
'    If CheckStr(adoRecordset.Fields("cu111")) = "Y" Then
'        Label3.ForeColor = &HFF&
'    Else
'        Label3.ForeColor = &H80000012
'    End If
'End If
'CheckOC
'
'If bolIsL = False Then
'        'Added by Lydia 2019/12/26 利益衝突案件：於後面增加欄位
'        SeColTM = " ,' ' as cnt,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno "
'        SeColPA = " ,' ' as cnt,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno "
'        SeColSP = " ,' ' as cnt,sp08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno "
'        SeColLC = " ,' ' as cnt,lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno "
'        SeColHC = " ,' ' as cnt,hc05 as cust01,hc24 as cust02,hc25 as cust03,hc26 as cust04,hc27 as cust05,'' as fcno "
'        'end 2019/12/26
'
'            'Modified by Morgan 2019/1/30 SQLGrpStr(Str02, 2)->GetAddStr(str02) 取消不必要系統別的檢查(減少語法執行次數,分所有較大影響)
'            'Modified by Lydia 2019/12/26 +增加欄位SeColTM
'            strSql = "SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') AS 本所案號 ,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁, NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1, NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度," & _
'                     "NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort " & _
'                           " FROM TRADEMARK,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress WHERE tm10=na01(+) and TM23='" & Me.Tag & "' AND SUBSTR(TM23,1,8) = c1.CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1)) = c1.CU02(+) " & SeColTM & _
'               " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
'               " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) " & strSQL1 & strSQL5
'            strSql = "SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') AS 本所案號 ,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁, NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1, NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度," & _
'                     "NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort " & SeColTM & _
'                           " FROM TRADEMARK,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress WHERE tm10=na01(+) and TM78='" & Me.Tag & "' AND SUBSTR(TM23,1,8) = c1.CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1)) = c1.CU02(+) " & _
'               " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
'               " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) " & strSQL1 & strSQL5
'            strSql = "SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') AS 本所案號 ,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁, NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1, NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度," & _
'                     "NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort " & SeColTM & _
'                           " FROM TRADEMARK,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress WHERE tm10=na01(+) and TM79='" & Me.Tag & "' AND SUBSTR(TM23,1,8) = c1.CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1)) = c1.CU02(+) " & _
'               " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
'               " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) " & strSQL1 & strSQL5
'            strSql = "SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') AS 本所案號 ,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁, NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1, NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度," & _
'                     "NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort " & SeColTM & _
'                           " FROM TRADEMARK,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress WHERE tm10=na01(+) and TM80='" & Me.Tag & "' AND SUBSTR(TM23,1,8) = c1.CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1)) = c1.CU02(+) " & _
'               " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
'               " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) " & strSQL1 & strSQL5
'            strSql = "SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') AS 本所案號 ,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁, NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1, NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度," & _
'                     "NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort " & SeColTM & _
'                           " FROM TRADEMARK,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress WHERE tm10=na01(+) and TM81='" & Me.Tag & "' AND SUBSTR(TM23,1,8) = c1.CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1)) = c1.CU02(+) " & _
'               " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
'               " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) " & strSQL1 & strSQL5
'
'            'Modify By Sindy 2014/7/7 +||DECODE(PA165,'Y','＃','')
'            'Modified by Lydia 2019/12/26 +增加欄位SeColPA
'            strSql = strSql + " union select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','') AS 本所案號 ,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , SUBSTRB(NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)),1,10) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'                     "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort " & SeColPA & _
'                     " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress WHERE pa09=na01(+) and (PA26='" & Me.Tag & "') and substr(pa26,1,8)=c1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=c1.cu02(+) " & _
'                     " and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'                     " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) " & strSQL2 & strSQL5 & "CUFA_PA"
'            strSql = strSql + " union select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','') AS 本所案號 ,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , SUBSTRB(NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)),1,10) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'                     "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort " & SeColPA & _
'                     " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress WHERE pa09=na01(+) and (PA27='" & Me.Tag & "') and substr(pa26,1,8)=c1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=c1.cu02(+) " & _
'                     " and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'                     " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) " & strSQL2 & strSQL5 & "CUFA_PA"
'            strSql = strSql + " union select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','') AS 本所案號 ,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , SUBSTRB(NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)),1,10) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'                     "DECODE(PA25,NULL,'','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort " & SeColPA & _
'                     " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress WHERE pa09=na01(+) and (PA28='" & Me.Tag & "') and substr(pa26,1,8)=c1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=c1.cu02(+) " & _
'                     " and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'                     " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) " & strSQL2 & strSQL5 & "CUFA_PA"
'            strSql = strSql + " union select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','') AS 本所案號 ,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , SUBSTRB(NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)),1,10) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'                     "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort " & SeColPA & _
'                     " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress WHERE pa09=na01(+) and (PA29='" & Me.Tag & "') and substr(pa26,1,8)=c1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=c1.cu02(+) " & _
'                     " and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'                     " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and Pa01=cp01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) " & strSQL2 & strSQL5 & "CUFA_PA"
'            strSql = strSql + " union select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','') AS 本所案號 ,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , SUBSTRB(NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)),1,10) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'                     "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort " & SeColPA & _
'                     " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress WHERE pa09=na01(+) and (PA30='" & Me.Tag & "') and substr(pa26,1,8)=c1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=c1.cu02(+) " & _
'                     " and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'                     " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) " & strSQL2 & strSQL5 & "CUFA_PA"
'
'            'Modified by Lydia 2019/12/26 +增加欄位SeColSP
'            'Modify by Amy 2020/02/05 +SP73 商品類別
'            strSql = strSql + " union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'                     "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'                     ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort " & SeColSP & _
'                     " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,caseprogress WHERE sp09=na01(+) and (SP08='" & Me.Tag & "') AND SUBSTR(SP08,1,8)=C1.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C1.CU02(+) AND SUBSTr(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'                     " and substr(sp65,1,8)=c4.cu01(+) and decode(substr(sp65,9,1),null,'0',substr(sp65,9,1))=c4.cu02(+) and substr(sp66,1,8)=c5.cu01(+) and decode(substr(sp66,9,1),null,'0',substr(sp66,9,1))=c5.cu02(+) and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) " & StrSQL3 & strSQL5 & "CUFA_SP"
'            strSql = strSql + " union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'                     "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'                     ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort " & SeColSP & _
'                     " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,caseprogress WHERE sp09=na01(+) and (SP58='" & Me.Tag & "') AND SUBSTR(SP08,1,8)=C1.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C1.CU02(+) AND SUBSTr(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'                     " and substr(sp65,1,8)=c4.cu01(+) and decode(substr(sp65,9,1),null,'0',substr(sp65,9,1))=c4.cu02(+) and substr(sp66,1,8)=c5.cu01(+) and decode(substr(sp66,9,1),null,'0',substr(sp66,9,1))=c5.cu02(+) and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) " & StrSQL3 & strSQL5 & "CUFA_SP"
'            strSql = strSql + " union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'                     "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'                     ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort " & SeColSP & _
'                     " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,caseprogress WHERE sp09=na01(+) and (SP59='" & Me.Tag & "') AND SUBSTR(SP08,1,8)=C1.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C1.CU02(+) AND SUBSTr(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'                     " and substr(sp65,1,8)=c4.cu01(+) and decode(substr(sp65,9,1),null,'0',substr(sp65,9,1))=c4.cu02(+) and substr(sp66,1,8)=c5.cu01(+) and decode(substr(sp66,9,1),null,'0',substr(sp66,9,1))=c5.cu02(+) and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) " & StrSQL3 & strSQL5 & "CUFA_SP"
'            strSql = strSql + " union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'                     "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'                     ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort " & SeColSP & _
'                     " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,caseprogress WHERE sp09=na01(+) and (SP65='" & Me.Tag & "') AND SUBSTR(SP08,1,8)=C1.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C1.CU02(+) AND SUBSTr(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'                     " and substr(sp65,1,8)=c4.cu01(+) and decode(substr(sp65,9,1),null,'0',substr(sp65,9,1))=c4.cu02(+) and substr(sp66,1,8)=c5.cu01(+) and decode(substr(sp66,9,1),null,'0',substr(sp66,9,1))=c5.cu02(+) and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) " & StrSQL3 & strSQL5 & "CUFA_SP"
'            strSql = strSql + " union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'                     "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'                     ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort " & SeColSP & _
'                     " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,caseprogress WHERE sp09=na01(+) and (SP66='" & Me.Tag & "') AND SUBSTR(SP08,1,8)=C1.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C1.CU02(+) AND SUBSTr(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'                     " and substr(sp65,1,8)=c4.cu01(+) and decode(substr(sp65,9,1),null,'0',substr(sp65,9,1))=c4.cu02(+) and substr(sp66,1,8)=c5.cu01(+) and decode(substr(sp66,9,1),null,'0',substr(sp66,9,1))=c5.cu02(+) and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) " & StrSQL3 & strSQL5 & "CUFA_SP"
'            'end 2020/02/05
'
'            'Modify By Sindy 2011/1/20 +LC43,LC44,LC45,LC46
'            'Modified by Lydia 2019/12/26 +增加欄位SeColLC
'            strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'                     ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
'                     " FROM LAWCASE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,caseprogress " & _
'                     " WHERE lc15=na01(+) and LC11='" & Me.Tag & "' " & _
'                     " and SUBSTR(LC11,1,8)=c1.CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = c1.CU02(+) " & _
'                     " and SUBSTR(LC43,1,8)=c2.CU01(+) AND DECODE(SUBSTR(LC43,9,1),NULL,'0',SUBSTR(LC43,9,1)) = c2.CU02(+) " & _
'                     " and SUBSTR(LC44,1,8)=c3.CU01(+) AND DECODE(SUBSTR(LC44,9,1),NULL,'0',SUBSTR(LC44,9,1)) = c3.CU02(+) " & _
'                     " and SUBSTR(LC45,1,8)=c4.CU01(+) AND DECODE(SUBSTR(LC45,9,1),NULL,'0',SUBSTR(LC45,9,1)) = c4.CU02(+) " & _
'                     " and SUBSTR(LC46,1,8)=c5.CU01(+) AND DECODE(SUBSTR(LC46,9,1),NULL,'0',SUBSTR(LC46,9,1)) = c5.CU02(+) " & _
'                     " and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+) " & StrSQL4 & strSQL5
'            strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'                     ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
'                     " FROM LAWCASE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,caseprogress " & _
'                     " WHERE lc15=na01(+) and LC43='" & Me.Tag & "' " & _
'                     " and SUBSTR(LC11,1,8)=c1.CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = c1.CU02(+) " & _
'                     " and SUBSTR(LC43,1,8)=c2.CU01(+) AND DECODE(SUBSTR(LC43,9,1),NULL,'0',SUBSTR(LC43,9,1)) = c2.CU02(+) " & _
'                     " and SUBSTR(LC44,1,8)=c3.CU01(+) AND DECODE(SUBSTR(LC44,9,1),NULL,'0',SUBSTR(LC44,9,1)) = c3.CU02(+) " & _
'                     " and SUBSTR(LC45,1,8)=c4.CU01(+) AND DECODE(SUBSTR(LC45,9,1),NULL,'0',SUBSTR(LC45,9,1)) = c4.CU02(+) " & _
'                     " and SUBSTR(LC46,1,8)=c5.CU01(+) AND DECODE(SUBSTR(LC46,9,1),NULL,'0',SUBSTR(LC46,9,1)) = c5.CU02(+) " & _
'                     " and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+) " & StrSQL4 & strSQL5
'            strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'                     ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
'                     " FROM LAWCASE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,caseprogress " & _
'                     " WHERE lc15=na01(+) and LC44='" & Me.Tag & "' " & _
'                     " and SUBSTR(LC11,1,8)=c1.CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = c1.CU02(+) " & _
'                     " and SUBSTR(LC43,1,8)=c2.CU01(+) AND DECODE(SUBSTR(LC43,9,1),NULL,'0',SUBSTR(LC43,9,1)) = c2.CU02(+) " & _
'                     " and SUBSTR(LC44,1,8)=c3.CU01(+) AND DECODE(SUBSTR(LC44,9,1),NULL,'0',SUBSTR(LC44,9,1)) = c3.CU02(+) " & _
'                     " and SUBSTR(LC45,1,8)=c4.CU01(+) AND DECODE(SUBSTR(LC45,9,1),NULL,'0',SUBSTR(LC45,9,1)) = c4.CU02(+) " & _
'                     " and SUBSTR(LC46,1,8)=c5.CU01(+) AND DECODE(SUBSTR(LC46,9,1),NULL,'0',SUBSTR(LC46,9,1)) = c5.CU02(+) " & _
'                     " and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+) " & StrSQL4 & strSQL5
'            strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'                     ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
'                     " FROM LAWCASE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,caseprogress " & _
'                     " WHERE lc15=na01(+) and LC45='" & Me.Tag & "' " & _
'                     " and SUBSTR(LC11,1,8)=c1.CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = c1.CU02(+) " & _
'                     " and SUBSTR(LC43,1,8)=c2.CU01(+) AND DECODE(SUBSTR(LC43,9,1),NULL,'0',SUBSTR(LC43,9,1)) = c2.CU02(+) " & _
'                     " and SUBSTR(LC44,1,8)=c3.CU01(+) AND DECODE(SUBSTR(LC44,9,1),NULL,'0',SUBSTR(LC44,9,1)) = c3.CU02(+) " & _
'                     " and SUBSTR(LC45,1,8)=c4.CU01(+) AND DECODE(SUBSTR(LC45,9,1),NULL,'0',SUBSTR(LC45,9,1)) = c4.CU02(+) " & _
'                     " and SUBSTR(LC46,1,8)=c5.CU01(+) AND DECODE(SUBSTR(LC46,9,1),NULL,'0',SUBSTR(LC46,9,1)) = c5.CU02(+) " & _
'                     " and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+) " & StrSQL4 & strSQL5
'            strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'                     ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
'                     " FROM LAWCASE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,caseprogress " & _
'                     " WHERE lc15=na01(+) and LC46='" & Me.Tag & "' " & _
'                     " and SUBSTR(LC11,1,8)=c1.CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = c1.CU02(+) " & _
'                     " and SUBSTR(LC43,1,8)=c2.CU01(+) AND DECODE(SUBSTR(LC43,9,1),NULL,'0',SUBSTR(LC43,9,1)) = c2.CU02(+) " & _
'                     " and SUBSTR(LC44,1,8)=c3.CU01(+) AND DECODE(SUBSTR(LC44,9,1),NULL,'0',SUBSTR(LC44,9,1)) = c3.CU02(+) " & _
'                     " and SUBSTR(LC45,1,8)=c4.CU01(+) AND DECODE(SUBSTR(LC45,9,1),NULL,'0',SUBSTR(LC45,9,1)) = c4.CU02(+) " & _
'                     " and SUBSTR(LC46,1,8)=c5.CU01(+) AND DECODE(SUBSTR(LC46,9,1),NULL,'0',SUBSTR(LC46,9,1)) = c5.CU02(+) " & _
'                     " and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+) " & StrSQL4 & strSQL5
'
'            'Modify By Sindy 2011/1/20 +HC24,HC25,HC26,HC27
'            'Modified by Lydia 2019/12/26 +增加欄位SeColHC
'            strSql = strSql + " union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,' ' AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'                     ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') as FSort " & SeColHC & _
'                     " FROM HIRECASE,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,caseprogress " & _
'                     " WHERE HC05='" & Me.Tag & "' " & _
'                     " AND SUBSTR(HC05,1,8)=c1.CU01(+) AND DECODE(SUBSTR(HC05,9,1),NULL,'0',SUBSTR(HC05,9,1))=c1.CU02(+) " & _
'                     " AND SUBSTR(HC24,1,8)=c2.CU01(+) AND DECODE(SUBSTR(HC24,9,1),NULL,'0',SUBSTR(HC24,9,1))=c2.CU02(+) " & _
'                     " AND SUBSTR(HC25,1,8)=c3.CU01(+) AND DECODE(SUBSTR(HC25,9,1),NULL,'0',SUBSTR(HC25,9,1))=c3.CU02(+) " & _
'                     " AND SUBSTR(HC26,1,8)=c4.CU01(+) AND DECODE(SUBSTR(HC26,9,1),NULL,'0',SUBSTR(HC26,9,1))=c4.CU02(+) " & _
'                     " AND SUBSTR(HC27,1,8)=c5.CU01(+) AND DECODE(SUBSTR(HC27,9,1),NULL,'0',SUBSTR(HC27,9,1))=c5.CU02(+) " & _
'                     " and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) " & strSQL8 & strSQL5
'            strSql = strSql + " union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,' ' AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'                     ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') as FSort " & SeColHC & _
'                     " FROM HIRECASE,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,caseprogress " & _
'                     " WHERE HC24='" & Me.Tag & "' " & _
'                     " AND SUBSTR(HC05,1,8)=c1.CU01(+) AND DECODE(SUBSTR(HC05,9,1),NULL,'0',SUBSTR(HC05,9,1))=c1.CU02(+) " & _
'                     " AND SUBSTR(HC24,1,8)=c2.CU01(+) AND DECODE(SUBSTR(HC24,9,1),NULL,'0',SUBSTR(HC24,9,1))=c2.CU02(+) " & _
'                     " AND SUBSTR(HC25,1,8)=c3.CU01(+) AND DECODE(SUBSTR(HC25,9,1),NULL,'0',SUBSTR(HC25,9,1))=c3.CU02(+) " & _
'                     " AND SUBSTR(HC26,1,8)=c4.CU01(+) AND DECODE(SUBSTR(HC26,9,1),NULL,'0',SUBSTR(HC26,9,1))=c4.CU02(+) " & _
'                     " AND SUBSTR(HC27,1,8)=c5.CU01(+) AND DECODE(SUBSTR(HC27,9,1),NULL,'0',SUBSTR(HC27,9,1))=c5.CU02(+) " & _
'                     " and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) " & strSQL8 & strSQL5
'            strSql = strSql + " union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,' ' AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'                     ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') as FSort  FROM HIRECASE,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,caseprogress " & _
'                     " WHERE HC25='" & Me.Tag & "' " & _
'                     " AND SUBSTR(HC05,1,8)=c1.CU01(+) AND DECODE(SUBSTR(HC05,9,1),NULL,'0',SUBSTR(HC05,9,1))=c1.CU02(+) " & _
'                     " AND SUBSTR(HC24,1,8)=c2.CU01(+) AND DECODE(SUBSTR(HC24,9,1),NULL,'0',SUBSTR(HC24,9,1))=c2.CU02(+) " & _
'                     " AND SUBSTR(HC25,1,8)=c3.CU01(+) AND DECODE(SUBSTR(HC25,9,1),NULL,'0',SUBSTR(HC25,9,1))=c3.CU02(+) " & _
'                     " AND SUBSTR(HC26,1,8)=c4.CU01(+) AND DECODE(SUBSTR(HC26,9,1),NULL,'0',SUBSTR(HC26,9,1))=c4.CU02(+) " & _
'                     " AND SUBSTR(HC27,1,8)=c5.CU01(+) AND DECODE(SUBSTR(HC27,9,1),NULL,'0',SUBSTR(HC27,9,1))=c5.CU02(+) " & _
'                     " and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) " & strSQL8 & strSQL5
'            strSql = strSql + " union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,' ' AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'                     ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') as FSort " & SeColHC & _
'                     " FROM HIRECASE,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,caseprogress " & _
'                     " WHERE HC26='" & Me.Tag & "' " & _
'                     " AND SUBSTR(HC05,1,8)=c1.CU01(+) AND DECODE(SUBSTR(HC05,9,1),NULL,'0',SUBSTR(HC05,9,1))=c1.CU02(+) " & _
'                     " AND SUBSTR(HC24,1,8)=c2.CU01(+) AND DECODE(SUBSTR(HC24,9,1),NULL,'0',SUBSTR(HC24,9,1))=c2.CU02(+) " & _
'                     " AND SUBSTR(HC25,1,8)=c3.CU01(+) AND DECODE(SUBSTR(HC25,9,1),NULL,'0',SUBSTR(HC25,9,1))=c3.CU02(+) " & _
'                     " AND SUBSTR(HC26,1,8)=c4.CU01(+) AND DECODE(SUBSTR(HC26,9,1),NULL,'0',SUBSTR(HC26,9,1))=c4.CU02(+) " & _
'                     " AND SUBSTR(HC27,1,8)=c5.CU01(+) AND DECODE(SUBSTR(HC27,9,1),NULL,'0',SUBSTR(HC27,9,1))=c5.CU02(+) " & _
'                     " and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) " & strSQL8 & strSQL5
'            strSql = strSql + " union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,' ' AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'                     ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') as FSort " & SeColHC & _
'                     " FROM HIRECASE,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,caseprogress " & _
'                     " WHERE HC27='" & Me.Tag & "' " & _
'                     " AND SUBSTR(HC05,1,8)=c1.CU01(+) AND DECODE(SUBSTR(HC05,9,1),NULL,'0',SUBSTR(HC05,9,1))=c1.CU02(+) " & _
'                     " AND SUBSTR(HC24,1,8)=c2.CU01(+) AND DECODE(SUBSTR(HC24,9,1),NULL,'0',SUBSTR(HC24,9,1))=c2.CU02(+) " & _
'                     " AND SUBSTR(HC25,1,8)=c3.CU01(+) AND DECODE(SUBSTR(HC25,9,1),NULL,'0',SUBSTR(HC25,9,1))=c3.CU02(+) " & _
'                     " AND SUBSTR(HC26,1,8)=c4.CU01(+) AND DECODE(SUBSTR(HC26,9,1),NULL,'0',SUBSTR(HC26,9,1))=c4.CU02(+) " & _
'                     " AND SUBSTR(HC27,1,8)=c5.CU01(+) AND DECODE(SUBSTR(HC27,9,1),NULL,'0',SUBSTR(HC27,9,1))=c5.CU02(+) " & _
'                     " and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) " & strSQL8 & strSQL5
'
'            '加考慮案件進度檔CP55, CP56, CP72欄位
'            'Modified by Lydia 2019/12/26 +增加欄位SeColTM
'            strSql = strSql + " union SELECT ' ' AS V,'△'||decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') AS 本所案號 ,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'                     ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort " & SeColTM & _
'                     " FROM TRADEMARK,nation,CUSTOMER C1, CUSTOMER C2,customer c3,customer c4,customer c5,caseprogress WHERE tm10=na01(+) and CP55='" & Me.Tag & "' And ((TM23<>'" & Me.Tag & "' Or TM23 Is Null) and (TM78<>'" & Me.Tag & "' Or TM78 Is Null) and (TM79<>'" & Me.Tag & "' Or TM79 Is Null) and (TM80<>'" & Me.Tag & "' Or TM80 Is Null) and (TM81<>'" & Me.Tag & "' Or TM81 Is Null)) and tm01 in (" & GetAddStr(Str02) & ") AND TM04='00' AND SUBSTR(TM23,1,8)=C1.CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1))=C1.CU02(+)" & _
'                     " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
'                     " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & strSQL1 & strSQL5
'            strSql = strSql + " union SELECT ' ' AS V,'△'||decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') AS 本所案號 ,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'                     ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort " & SeColTM & _
'                     " FROM TRADEMARK,nation,CUSTOMER C1, CUSTOMER C2,customer c3,customer c4,customer c5,caseprogress WHERE tm10=na01(+) and CP56='" & Me.Tag & "' And ((TM23<>'" & Me.Tag & "' Or TM23 Is Null) and (TM78<>'" & Me.Tag & "' Or TM78 Is Null) and (TM79<>'" & Me.Tag & "' Or TM79 Is Null) and (TM80<>'" & Me.Tag & "' Or TM80 Is Null) and (TM81<>'" & Me.Tag & "' Or TM81 Is Null)) and tm01 in (" & GetAddStr(Str02) & ") AND TM04='00' AND SUBSTR(TM23,1,8)=C1.CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1))=C1.CU02(+)" & _
'                     " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
'                     " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & strSQL1 & strSQL5
'            strSql = strSql + " union SELECT ' ' AS V,'△'||decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') AS 本所案號 ,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'                     ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort " & SeColTM & _
'                     " FROM TRADEMARK,nation,CUSTOMER C1, CUSTOMER C2,customer c3,customer c4,customer c5,caseprogress WHERE tm10=na01(+) and CP72='" & Me.Tag & "' And ((TM23<>'" & Me.Tag & "' Or TM23 Is Null) and (TM78<>'" & Me.Tag & "' Or TM78 Is Null) and (TM79<>'" & Me.Tag & "' Or TM79 Is Null) and (TM80<>'" & Me.Tag & "' Or TM80 Is Null) and (TM81<>'" & Me.Tag & "' Or TM81 Is Null)) and tm01 in (" & GetAddStr(Str02) & ") AND TM04='00' AND SUBSTR(TM23,1,8)=C1.CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1))=C1.CU02(+)" & _
'                     " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
'                     " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) " & strSQL1 & strSQL5
'
'            'Modify By Sindy 2014/7/7 +||DECODE(PA165,'Y','＃','')
'            'Modified by Lydia 2019/12/26 +增加欄位SeColPA
'            strSql = strSql + " union select ' ' AS V,'△'||decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','') AS 本所案號 ,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , SUBSTRB(NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)),1,10) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'                     "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度" & _
'                     ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort " & SeColPA & _
'                     " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5, CUSTOMER C6,caseprogress WHERE pa09=na01(+) and (CP55='" & Me.Tag & "') And ((PA26<>'" & Me.Tag & "' Or PA26 Is Null) And (PA27='" & Me.Tag & "' Or PA27 Is Null) And (PA28<>'" & Me.Tag & "' Or PA28 Is Null) And (PA29='" & Me.Tag & "' Or PA29 Is Null) And (PA30<>'" & Me.Tag & "' Or PA30 Is Null)) and pa01 in (" & GetAddStr(Str02) & ") and pa04='00' and substr(CP55,1,8)=c1.cu01(+) and decode(substr(CP55,9,1),null,'0',substr(CP55,9,1))=c1.cu02(+) " & _
'                     " AND SUBSTR(PA26,1,8)=C6.CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=C6.CU02(+) and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'                     " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & strSQL2 & strSQL5 & "CUFA_PA"
'            strSql = strSql + " union select ' ' AS V,'△'||decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','') AS 本所案號 ,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , SUBSTRB(NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)),1,10) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'                     "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度" & _
'                     ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort " & SeColPA & _
'                     " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5, CUSTOMER C6,caseprogress WHERE pa09=na01(+) and (CP56='" & Me.Tag & "') And ((PA26<>'" & Me.Tag & "' Or PA26 Is Null) And (PA27='" & Me.Tag & "' Or PA27 Is Null) And (PA28<>'" & Me.Tag & "' Or PA28 Is Null) And (PA29='" & Me.Tag & "' Or PA29 Is Null) And (PA30<>'" & Me.Tag & "' Or PA30 Is Null)) and pa01 in (" & GetAddStr(Str02) & ") and pa04='00' and substr(CP56,1,8)=c1.cu01(+) and decode(substr(CP56,9,1),null,'0',substr(CP56,9,1))=c1.cu02(+) " & _
'                     " AND SUBSTR(PA26,1,8)=C6.CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=C6.CU02(+) and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'                     " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & strSQL2 & strSQL5 & "CUFA_PA"
'            strSql = strSql + " union select ' ' AS V,'△'||decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','') AS 本所案號 ,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , SUBSTRB(NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)),1,10) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
'                     "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度" & _
'                     ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort " & SeColPA & _
'                     " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5, CUSTOMER C6,caseprogress WHERE pa09=na01(+) and (CP72='" & Me.Tag & "') And ((PA26<>'" & Me.Tag & "' Or PA26 Is Null) And (PA27='" & Me.Tag & "' Or PA27 Is Null) And (PA28<>'" & Me.Tag & "' Or PA28 Is Null) And (PA29='" & Me.Tag & "' Or PA29 Is Null) And (PA30<>'" & Me.Tag & "' Or PA30 Is Null)) and pa01 in (" & GetAddStr(Str02) & ") and pa04='00' and substr(CP72,1,8)=c1.cu01(+) and decode(substr(CP72,9,1),null,'0',substr(CP72,9,1))=c1.cu02(+) " & _
'                     " AND SUBSTR(PA26,1,8)=C6.CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=C6.CU02(+) and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
'                     " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & strSQL2 & strSQL5 & "CUFA_PA"
'
'            'Modify By Sindy 2011/1/20 整理SQL
'            'Modified by Lydia 2019/12/26 +增加欄位SeColSP
'            'Modify by Amy 2020/02/05 +SP73 商品類別
'            strSql = strSql + " union select ' ' AS V,'△'||SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , SUBSTRB(NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)),1,10) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'                     "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'                     ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort " & SeColSP & _
'                     " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,customer c5, CUSTOMER C6,caseprogress " & _
'                     " WHERE sp09=na01(+) and (CP55='" & Me.Tag & "') And ((SP08<>'" & Me.Tag & "' Or SP08 Is Null) And (SP58<>'" & Me.Tag & "' Or SP58 Is Null) And (SP59<>'" & Me.Tag & "' Or SP59 Is Null)) and sp01 in (" & GetAddStr(Str02) & ") AND SP04='00' " & _
'                     " AND SUBSTR(CP55,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP55,9,1),NULL,'0',SUBSTR(CP55,9,1))=C1.CU02(+) " & _
'                     " AND SUBSTR(SP08,1,8)=C6.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C6.CU02(+) " & _
'                     " AND SUBSTR(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) " & _
'                     " AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'                     " AND SUBSTR(SP65,1,8)=C4.CU01(+) AND DECODE(SUBSTR(SP65,9,1),NULL,'0',SUBSTR(SP65,9,1))=C4.CU02(+) " & _
'                     " AND SUBSTR(SP66,1,8)=C5.CU01(+) AND DECODE(SUBSTR(SP66,9,1),NULL,'0',SUBSTR(SP66,9,1))=C5.CU02(+) " & _
'                     " and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & StrSQL3 & strSQL5 & "CUFA_SP"
'            strSql = strSql + " union select ' ' AS V,'△'||SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , SUBSTRB(NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)),1,10) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'                     "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'                     ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort " & SeColSP & _
'                     " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,customer c5, CUSTOMER C6,caseprogress " & _
'                     " WHERE sp09=na01(+) and (CP56='" & Me.Tag & "') And ((SP08<>'" & Me.Tag & "' Or SP08 Is Null) And (SP58<>'" & Me.Tag & "' Or SP58 Is Null) And (SP59<>'" & Me.Tag & "' Or SP59 Is Null)) and sp01 in (" & GetAddStr(Str02) & ") AND SP04='00' " & _
'                     " AND SUBSTR(CP56,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP56,9,1),NULL,'0',SUBSTR(CP56,9,1))=C1.CU02(+) " & _
'                     " AND SUBSTR(SP08,1,8)=C6.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C6.CU02(+) " & _
'                     " AND SUBSTR(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) " & _
'                     " AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'                     " AND SUBSTR(SP65,1,8)=C4.CU01(+) AND DECODE(SUBSTR(SP65,9,1),NULL,'0',SUBSTR(SP65,9,1))=C4.CU02(+) " & _
'                     " AND SUBSTR(SP66,1,8)=C5.CU01(+) AND DECODE(SUBSTR(SP66,9,1),NULL,'0',SUBSTR(SP66,9,1))=C5.CU02(+) " & _
'                     " and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & StrSQL3 & strSQL5 & "CUFA_SP"
'            strSql = strSql + " union select ' ' AS V,'△'||SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , SUBSTRB(NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)),1,10) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
'                     "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
'                     ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort " & SeColSP & _
'                     " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,customer c5, CUSTOMER C6,caseprogress " & _
'                     " WHERE sp09=na01(+) and (CP72='" & Me.Tag & "') And ((SP08<>'" & Me.Tag & "' Or SP08 Is Null) And (SP58<>'" & Me.Tag & "' Or SP58 Is Null) And (SP59<>'" & Me.Tag & "' Or SP59 Is Null)) and sp01 in (" & GetAddStr(Str02) & ") AND SP04='00' " & _
'                     " AND SUBSTR(CP72,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP72,9,1),NULL,'0',SUBSTR(CP72,9,1))=C1.CU02(+) " & _
'                     " AND SUBSTR(SP08,1,8)=C6.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C6.CU02(+) " & _
'                     " AND SUBSTR(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) " & _
'                     " AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
'                     " AND SUBSTR(SP65,1,8)=C4.CU01(+) AND DECODE(SUBSTR(SP65,9,1),NULL,'0',SUBSTR(SP65,9,1))=C4.CU02(+) " & _
'                     " AND SUBSTR(SP66,1,8)=C5.CU01(+) AND DECODE(SUBSTR(SP66,9,1),NULL,'0',SUBSTR(SP66,9,1))=C5.CU02(+) " & _
'                     " and CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & StrSQL3 & strSQL5 & "CUFA_SP"
'            'end 2020/02/05
'
'            'Modify By Sindy 2011/1/20 +LC43,LC44,LC45,LC46
'            'Modified by Lydia 2019/12/26 +增加欄位SeColLC
'            strSql = strSql + " union select ' ' AS V,'△'||LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , SUBSTRB(NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)),1,10) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
'                     " FROM LAWCASE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,customer c5, CUSTOMER C6,caseprogress " & _
'                     " WHERE lc15=na01(+) and CP55='" & Me.Tag & "' And (LC11<>'" & Me.Tag & "' Or LC11 Is Null) And (LC43<>'" & Me.Tag & "' Or LC43 Is Null) And (LC44<>'" & Me.Tag & "' Or LC44 Is Null) And (LC45<>'" & Me.Tag & "' Or LC45 Is Null) And (LC46<>'" & Me.Tag & "' Or LC46 Is Null) and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00' " & _
'                     " AND SUBSTR(CP55,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP55,9,1),NULL,'0',SUBSTR(CP55,9,1)) = C1.CU02(+) " & _
'                     " AND SUBSTR(LC11,1,8)=C6.CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1))=C6.CU02(+) " & _
'                     " AND SUBSTR(LC43,1,8)=C2.CU01(+) AND DECODE(SUBSTR(LC43,9,1),NULL,'0',SUBSTR(LC43,9,1))=C2.CU02(+) " & _
'                     " AND SUBSTR(LC44,1,8)=C3.CU01(+) AND DECODE(SUBSTR(LC44,9,1),NULL,'0',SUBSTR(LC44,9,1))=C3.CU02(+) " & _
'                     " AND SUBSTR(LC45,1,8)=C4.CU01(+) AND DECODE(SUBSTR(LC45,9,1),NULL,'0',SUBSTR(LC45,9,1))=C4.CU02(+) " & _
'                     " AND SUBSTR(LC46,1,8)=C5.CU01(+) AND DECODE(SUBSTR(LC46,9,1),NULL,'0',SUBSTR(LC46,9,1))=C5.CU02(+) " & _
'                     " and CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) " & StrSQL4 & strSQL5
'            strSql = strSql + " union select ' ' AS V,'△'||LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , SUBSTRB(NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)),1,10) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
'                     " FROM LAWCASE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,customer c5, CUSTOMER C6,caseprogress " & _
'                     " WHERE lc15=na01(+) and CP56='" & Me.Tag & "' And (LC11<>'" & Me.Tag & "' Or LC11 Is Null) And (LC43<>'" & Me.Tag & "' Or LC43 Is Null) And (LC44<>'" & Me.Tag & "' Or LC44 Is Null) And (LC45<>'" & Me.Tag & "' Or LC45 Is Null) And (LC46<>'" & Me.Tag & "' Or LC46 Is Null) and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00' " & _
'                     " AND SUBSTR(CP56,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP56,9,1),NULL,'0',SUBSTR(CP56,9,1)) = C1.CU02(+) " & _
'                     " AND SUBSTR(LC11,1,8)=C6.CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1))=C6.CU02(+) " & _
'                     " AND SUBSTR(LC43,1,8)=C2.CU01(+) AND DECODE(SUBSTR(LC43,9,1),NULL,'0',SUBSTR(LC43,9,1))=C2.CU02(+) " & _
'                     " AND SUBSTR(LC44,1,8)=C3.CU01(+) AND DECODE(SUBSTR(LC44,9,1),NULL,'0',SUBSTR(LC44,9,1))=C3.CU02(+) " & _
'                     " AND SUBSTR(LC45,1,8)=C4.CU01(+) AND DECODE(SUBSTR(LC45,9,1),NULL,'0',SUBSTR(LC45,9,1))=C4.CU02(+) " & _
'                     " AND SUBSTR(LC46,1,8)=C5.CU01(+) AND DECODE(SUBSTR(LC46,9,1),NULL,'0',SUBSTR(LC46,9,1))=C5.CU02(+) " & _
'                     " and CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) " & StrSQL4 & strSQL5
'            strSql = strSql + " union select ' ' AS V,'△'||LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , SUBSTRB(NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)),1,10) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
'                     " FROM LAWCASE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,customer c5, CUSTOMER C6,caseprogress " & _
'                     " WHERE lc15=na01(+) and CP72='" & Me.Tag & "' And (LC11<>'" & Me.Tag & "' Or LC11 Is Null) And (LC43<>'" & Me.Tag & "' Or LC43 Is Null) And (LC44<>'" & Me.Tag & "' Or LC44 Is Null) And (LC45<>'" & Me.Tag & "' Or LC45 Is Null) And (LC46<>'" & Me.Tag & "' Or LC46 Is Null) and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00' " & _
'                     " AND SUBSTR(CP72,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP72,9,1),NULL,'0',SUBSTR(CP72,9,1)) = C1.CU02(+) " & _
'                     " AND SUBSTR(LC11,1,8)=C6.CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1))=C6.CU02(+) " & _
'                     " AND SUBSTR(LC43,1,8)=C2.CU01(+) AND DECODE(SUBSTR(LC43,9,1),NULL,'0',SUBSTR(LC43,9,1))=C2.CU02(+) " & _
'                     " AND SUBSTR(LC44,1,8)=C3.CU01(+) AND DECODE(SUBSTR(LC44,9,1),NULL,'0',SUBSTR(LC44,9,1))=C3.CU02(+) " & _
'                     " AND SUBSTR(LC45,1,8)=C4.CU01(+) AND DECODE(SUBSTR(LC45,9,1),NULL,'0',SUBSTR(LC45,9,1))=C4.CU02(+) " & _
'                     " AND SUBSTR(LC46,1,8)=C5.CU01(+) AND DECODE(SUBSTR(LC46,9,1),NULL,'0',SUBSTR(LC46,9,1))=C5.CU02(+) " & _
'                     " and CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) " & StrSQL4 & strSQL5
'
'            'Modify By Sindy 2011/1/20 +HC24,HC25,HC26,HC27
'            'Modified by Lydia 2019/12/26 +增加欄位SeColHC
'            strSql = strSql + " union select ' ' AS V,'△'||HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,' ' AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , SUBSTRB(NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)),1,10) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') as FSort " & SeColHC & _
'                     " FROM HIRECASE,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,customer c5, CUSTOMER C6,caseprogress " & _
'                     " WHERE CP55='" & Me.Tag & "' And (HC05<>'" & Me.Tag & "' Or HC05 Is Null) And (HC24<>'" & Me.Tag & "' Or HC24 Is Null) And (HC25<>'" & Me.Tag & "' Or HC25 Is Null) And (HC26<>'" & Me.Tag & "' Or HC26 Is Null) And (HC27<>'" & Me.Tag & "' Or HC27 Is Null) and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00' " & _
'                     " AND SUBSTR(CP55,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP55,9,1),NULL,'0',SUBSTR(CP55,9,1))=C1.CU02(+) " & _
'                     " AND SUBSTR(HC05,1,8)=C6.CU01(+) AND DECODE(SUBSTR(HC05,9,1),NULL,'0',SUBSTR(HC05,9,1))=C6.CU02(+) " & _
'                     " AND SUBSTR(HC24,1,8)=C2.CU01(+) AND DECODE(SUBSTR(HC24,9,1),NULL,'0',SUBSTR(HC24,9,1))=C2.CU02(+) " & _
'                     " AND SUBSTR(HC25,1,8)=C3.CU01(+) AND DECODE(SUBSTR(HC25,9,1),NULL,'0',SUBSTR(HC25,9,1))=C3.CU02(+) " & _
'                     " AND SUBSTR(HC26,1,8)=C4.CU01(+) AND DECODE(SUBSTR(HC26,9,1),NULL,'0',SUBSTR(HC26,9,1))=C4.CU02(+) " & _
'                     " AND SUBSTR(HC27,1,8)=C5.CU01(+) AND DECODE(SUBSTR(HC27,9,1),NULL,'0',SUBSTR(HC27,9,1))=C5.CU02(+) " & _
'                     " and CP01=HC01(+) and CP02=HC02(+) and CP03=HC03(+) and CP04=HC04(+) " & strSQL8 & strSQL5
'            strSql = strSql + " union select ' ' AS V,'△'||HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,' ' AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , SUBSTRB(NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)),1,10) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') as FSort " & SeColHC & _
'                     " FROM HIRECASE,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,customer c5, CUSTOMER C6,caseprogress " & _
'                     " WHERE CP56='" & Me.Tag & "' And (HC05<>'" & Me.Tag & "' Or HC05 Is Null) And (HC24<>'" & Me.Tag & "' Or HC24 Is Null) And (HC25<>'" & Me.Tag & "' Or HC25 Is Null) And (HC26<>'" & Me.Tag & "' Or HC26 Is Null) And (HC27<>'" & Me.Tag & "' Or HC27 Is Null) and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00' " & _
'                     " AND SUBSTR(CP56,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP56,9,1),NULL,'0',SUBSTR(CP56,9,1))=C1.CU02(+) " & _
'                     " AND SUBSTR(HC05,1,8)=C6.CU01(+) AND DECODE(SUBSTR(HC05,9,1),NULL,'0',SUBSTR(HC05,9,1))=C6.CU02(+) " & _
'                     " AND SUBSTR(HC24,1,8)=C2.CU01(+) AND DECODE(SUBSTR(HC24,9,1),NULL,'0',SUBSTR(HC24,9,1))=C2.CU02(+) " & _
'                     " AND SUBSTR(HC25,1,8)=C3.CU01(+) AND DECODE(SUBSTR(HC25,9,1),NULL,'0',SUBSTR(HC25,9,1))=C3.CU02(+) " & _
'                     " AND SUBSTR(HC26,1,8)=C4.CU01(+) AND DECODE(SUBSTR(HC26,9,1),NULL,'0',SUBSTR(HC26,9,1))=C4.CU02(+) " & _
'                     " AND SUBSTR(HC27,1,8)=C5.CU01(+) AND DECODE(SUBSTR(HC27,9,1),NULL,'0',SUBSTR(HC27,9,1))=C5.CU02(+) " & _
'                     " and CP01=HC01(+) and CP02=HC02(+) and CP03=HC03(+) and CP04=HC04(+) " & strSQL8 & strSQL5
'            strSql = strSql + " union select ' ' AS V,'△'||HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,' ' AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , SUBSTRB(NVL(C6.CU04,NVL(C6.CU05||C6.CU88||C6.CU89||C6.CU90,C6.CU06)),1,10) AS 申請人1 ,' ' AS 商品類別,'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') as FSort " & SeColHC & _
'                     " FROM HIRECASE,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,customer c5, CUSTOMER C6,caseprogress " & _
'                     " WHERE CP72='" & Me.Tag & "' And (HC05<>'" & Me.Tag & "' Or HC05 Is Null) And (HC24<>'" & Me.Tag & "' Or HC24 Is Null) And (HC25<>'" & Me.Tag & "' Or HC25 Is Null) And (HC26<>'" & Me.Tag & "' Or HC26 Is Null) And (HC27<>'" & Me.Tag & "' Or HC27 Is Null) and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00' " & _
'                     " AND SUBSTR(CP72,1,8)=C1.CU01(+) AND DECODE(SUBSTR(CP72,9,1),NULL,'0',SUBSTR(CP72,9,1))=C1.CU02(+) " & _
'                     " AND SUBSTR(HC05,1,8)=C6.CU01(+) AND DECODE(SUBSTR(HC05,9,1),NULL,'0',SUBSTR(HC05,9,1))=C6.CU02(+) " & _
'                     " AND SUBSTR(HC24,1,8)=C2.CU01(+) AND DECODE(SUBSTR(HC24,9,1),NULL,'0',SUBSTR(HC24,9,1))=C2.CU02(+) " & _
'                     " AND SUBSTR(HC25,1,8)=C3.CU01(+) AND DECODE(SUBSTR(HC25,9,1),NULL,'0',SUBSTR(HC25,9,1))=C3.CU02(+) " & _
'                     " AND SUBSTR(HC26,1,8)=C4.CU01(+) AND DECODE(SUBSTR(HC26,9,1),NULL,'0',SUBSTR(HC26,9,1))=C4.CU02(+) " & _
'                     " AND SUBSTR(HC27,1,8)=C5.CU01(+) AND DECODE(SUBSTR(HC27,9,1),NULL,'0',SUBSTR(HC27,9,1))=C5.CU02(+) " & _
'                     " and CP01=HC01(+) and CP02=HC02(+) and CP03=HC03(+) and CP04=HC04(+) " & strSQL8 & strSQL5
'Else  '法務專用 add by nickc 2005/10/04
'
'        'Added by Lydia 2019/12/26 利益衝突案件：於後面增加欄位
'        SeColTM = " ,' ' as CP09,' ' as cnt,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno "
'        SeColPA = " ,' ' as CP09,' ' as cnt,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno "
'        SeColSP = " ,' ' as CP09,' ' as cnt,sp08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno "
'        SeColLC = " ,' ' as CP09,' ' as cnt,lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno "
'        SeColHC = " ,' ' as CP09,' ' as cnt,hc05 as cust01,hc24 as cust02,hc25 as cust03,hc26 as cust04,hc27 as cust05,'' as fcno "
'        'end 2019/12/26
'
'         'Modify By Sindy 2011/1/21 +LC43,LC44,LC45,LC46
'         'Modified by Lydia 2019/12/26 +增加欄位SeColLC
'         strSql = "select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家, sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
'                     " FROM LAWCASE,nation,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
'                     " WHERE lc15=na01(+) " & _
'                     " and LC11='" & Me.Tag & "' and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00' " & _
'                     " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1)) = CU02(+) " & _
'                     " and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+) and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & StrSQL4 & strSQL5
'         strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家, sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
'                     " FROM LAWCASE,nation,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
'                     " WHERE lc15=na01(+) " & _
'                     " and LC43='" & Me.Tag & "' and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00' " & _
'                     " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1)) = CU02(+) " & _
'                     " and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+) and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & StrSQL4 & strSQL5
'         strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家, sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
'                     " FROM LAWCASE,nation,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
'                     " WHERE lc15=na01(+) " & _
'                     " and LC44='" & Me.Tag & "' and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00' " & _
'                     " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1)) = CU02(+) " & _
'                     " and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+) and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & StrSQL4 & strSQL5
'         strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家, sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
'                     " FROM LAWCASE,nation,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
'                     " WHERE lc15=na01(+) " & _
'                     " and LC45='" & Me.Tag & "' and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00' " & _
'                     " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1)) = CU02(+) " & _
'                     " and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+) and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & StrSQL4 & strSQL5
'         strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家, sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
'                     " FROM LAWCASE,nation,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
'                     " WHERE lc15=na01(+) " & _
'                     " and LC46='" & Me.Tag & "' and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00' " & _
'                     " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1)) = CU02(+) " & _
'                     " and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+) and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & StrSQL4 & strSQL5
'
'         'Modify By Sindy 2011/1/21 +HC24,HC25,HC26,HC27
'         'Modified by Lydia 2019/12/26 +增加欄位SeColHC
'         strSql = strSql + " union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,decode(cp10,'0',sqldatet(cp53),'台灣')||'-'||decode(cp10,'0',sqldatet(cp54),'') AS 申請國家, sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') as FSort " & SeColHC & _
'                     " FROM HIRECASE,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
'                     " WHERE HC05='" & Me.Tag & "' and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00' " & _
'                     " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+) " & _
'                     " and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL8 & strSQL5
'         strSql = strSql + " union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,decode(cp10,'0',sqldatet(cp53),'台灣')||'-'||decode(cp10,'0',sqldatet(cp54),'') AS 申請國家, sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') as FSort " & SeColHC & _
'                     " FROM HIRECASE,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
'                     " WHERE HC24='" & Me.Tag & "' and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00' " & _
'                     " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+) " & _
'                     " and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL8 & strSQL5
'         strSql = strSql + " union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,decode(cp10,'0',sqldatet(cp53),'台灣')||'-'||decode(cp10,'0',sqldatet(cp54),'') AS 申請國家, sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') as FSort " & SeColHC & _
'                     " FROM HIRECASE,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
'                     " WHERE HC25='" & Me.Tag & "' and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00' " & _
'                     " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+) " & _
'                     " and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL8 & strSQL5
'         strSql = strSql + " union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,decode(cp10,'0',sqldatet(cp53),'台灣')||'-'||decode(cp10,'0',sqldatet(cp54),'') AS 申請國家, sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') as FSort " & SeColHC & _
'                     " FROM HIRECASE,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
'                     " WHERE HC26='" & Me.Tag & "' and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00' " & _
'                     " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+) " & _
'                     " and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL8 & strSQL5
'         strSql = strSql + " union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,decode(cp10,'0',sqldatet(cp53),'台灣')||'-'||decode(cp10,'0',sqldatet(cp54),'') AS 申請國家, sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') as FSort " & SeColHC & _
'                     " FROM HIRECASE,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
'                     " WHERE HC27='" & Me.Tag & "' and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00' " & _
'                     " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+) " & _
'                     " and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL8 & strSQL5
'
'         '加考慮案件進度檔CP55, CP56, CP72欄位
'         'Modify By Sindy 2011/1/21 +LC43,LC44,LC45,LC46
'         'Modified by Lydia 2019/12/26 +增加欄位SeColLC
'         strSql = strSql + " union select ' ' AS V,'△'||LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家, sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人," & _
'                     "decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
'                     " FROM LAWCASE,nation,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
'                     " WHERE lc15=na01(+) and CP55='" & Me.Tag & "' And (LC11<>'" & Me.Tag & "' Or LC11 Is Null) And (LC43<>'" & Me.Tag & "' Or LC43 Is Null) And (LC44<>'" & Me.Tag & "' Or LC44 Is Null) And (LC45<>'" & Me.Tag & "' Or LC45 Is Null) And (LC46<>'" & Me.Tag & "' Or LC46 Is Null) and cp01 in (" & GetAddStr(Str02) & ") AND cp04='00' " & _
'                     " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+) " & _
'                     " and CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) and lc01 is not null and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & StrSQL4 & strSQL5
'         strSql = strSql + " union select ' ' AS V,'△'||LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家, sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人," & _
'                     "decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
'                     " FROM LAWCASE,nation,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
'                     " WHERE lc15=na01(+) and CP56='" & Me.Tag & "' And (LC11<>'" & Me.Tag & "' Or LC11 Is Null) And (LC43<>'" & Me.Tag & "' Or LC43 Is Null) And (LC44<>'" & Me.Tag & "' Or LC44 Is Null) And (LC45<>'" & Me.Tag & "' Or LC45 Is Null) And (LC46<>'" & Me.Tag & "' Or LC46 Is Null) and cp01 in (" & GetAddStr(Str02) & ") AND cp04='00' " & _
'                     " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+) " & _
'                     " and CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) and lc01 is not null and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & StrSQL4 & strSQL5
'         strSql = strSql + " union select ' ' AS V,'△'||LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家, sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人," & _
'                     "decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
'                     " FROM LAWCASE,nation,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
'                     " WHERE lc15=na01(+) and CP72='" & Me.Tag & "' And (LC11<>'" & Me.Tag & "' Or LC11 Is Null) And (LC43<>'" & Me.Tag & "' Or LC43 Is Null) And (LC44<>'" & Me.Tag & "' Or LC44 Is Null) And (LC45<>'" & Me.Tag & "' Or LC45 Is Null) And (LC46<>'" & Me.Tag & "' Or LC46 Is Null) and cp01 in (" & GetAddStr(Str02) & ") AND cp04='00' " & _
'                     " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+) " & _
'                     " and CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) and lc01 is not null and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & StrSQL4 & strSQL5
'
'         'Modify By Sindy 2011/1/21 +HC24,HC25,HC26,HC27
'         'Modified by Lydia 2019/12/26 +增加欄位SeColHC
'         strSql = strSql + " union select ' ' AS V,'△'||HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,decode(cp10,'0',sqldatet(cp53),'台灣')||'-'||decode(cp10,'0',sqldatet(cp54),'') AS 申請國家, sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人," & _
'                     "decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort " & SeColHC & _
'                     " FROM HIRECASE,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
'                     " WHERE CP55='" & Me.Tag & "' And (HC05<>'" & Me.Tag & "' Or HC05 Is Null) And (HC24<>'" & Me.Tag & "' Or HC24 Is Null) And (HC25<>'" & Me.Tag & "' Or HC25 Is Null) And (HC26<>'" & Me.Tag & "' Or HC26 Is Null) And (HC27<>'" & Me.Tag & "' Or HC27 Is Null) and cp01 in (" & GetAddStr(Str02) & ") AND cp04='00' " & _
'                     " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+) " & _
'                     " and CP01=HC01(+) and CP02=HC02(+) and CP03=HC03(+) and CP04=HC04(+) and hc01 is not null and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL8 & strSQL5
'         strSql = strSql + " union select ' ' AS V,'△'||HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,decode(cp10,'0',sqldatet(cp53),'台灣')||'-'||decode(cp10,'0',sqldatet(cp54),'') AS 申請國家, sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人," & _
'                     "decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort " & SeColHC & _
'                     " FROM HIRECASE,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
'                     " WHERE CP56='" & Me.Tag & "' And (HC05<>'" & Me.Tag & "' Or HC05 Is Null) And (HC24<>'" & Me.Tag & "' Or HC24 Is Null) And (HC25<>'" & Me.Tag & "' Or HC25 Is Null) And (HC26<>'" & Me.Tag & "' Or HC26 Is Null) And (HC27<>'" & Me.Tag & "' Or HC27 Is Null) and cp01 in (" & GetAddStr(Str02) & ") AND cp04='00' " & _
'                     " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+) " & _
'                     " and CP01=HC01(+) and CP02=HC02(+) and CP03=HC03(+) and CP04=HC04(+) and hc01 is not null and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL8 & strSQL5
'         strSql = strSql + " union select ' ' AS V,'△'||HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,decode(cp10,'0',sqldatet(cp53),'台灣')||'-'||decode(cp10,'0',sqldatet(cp54),'') AS 申請國家, sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人," & _
'                     "decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort " & SeColHC & _
'                     " FROM HIRECASE,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
'                     " WHERE CP72='" & Me.Tag & "' And (HC05<>'" & Me.Tag & "' Or HC05 Is Null) And (HC24<>'" & Me.Tag & "' Or HC24 Is Null) And (HC25<>'" & Me.Tag & "' Or HC25 Is Null) And (HC26<>'" & Me.Tag & "' Or HC26 Is Null) And (HC27<>'" & Me.Tag & "' Or HC27 Is Null) and cp01 in (" & GetAddStr(Str02) & ") AND cp04='00' " & _
'                     " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+) " & _
'                     " and CP01=HC01(+) and CP02=HC02(+) and CP03=HC03(+) and CP04=HC04(+) and hc01 is not null and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL8 & strSQL5
'End If
'strSql = strSql & " ORDER BY FSort,本所案號 "
'
''Added by Lydia 2019/11/01 利益衝突案件：處理替換字串
''Mark by Lydia 2019/12/26
''If m_CuFaArea <> "" And stConPA & stConSP <> "" Then
''    stCuFaSQL = strSql
''    stCuFaSQL = Replace(stCuFaSQL, "CUFA_PA", stConPA)
''    stCuFaSQL = Replace(stCuFaSQL, "CUFA_SP", stConSP)
''    intI = 1
''    Set rsCnt = Nothing
''    Set rsCnt = ClsLawReadRstMsg(intI, stCuFaSQL)
''End If
''strSql = Replace(strSql, "CUFA_PA", "")
''strSql = Replace(strSql, "CUFA_SP", "")
'''end 2019/11/01
''end 2019/12/26
'
'CheckOC
'adoRecordset.CursorLocation = adUseClient
''Modified by Lydia 2019/12/26 改變型態
''adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
'
'If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'    If Len(Trim(Str02)) <> 0 Then
'        strTemp = Split(Str02, ",")
'    End If
'
'    'Modified by Lydia 2019/12/26 利益衝突案件：逐案號判斷
'    'adoRecordset.MoveFirst
'    'Dim StrTest2 As String, StrTest4 As String, s As Integer
'    'Set m_adoRst = adoRecordset.Clone 'Added by Lydia 2018/02/09 'move by Lydia 2018/12/17 從下面移上來
'    'If adoRecordset.RecordCount = 0 Then
'    '    Me.Enabled = True
'    '    cmdOK(0).Enabled = False
'    '    cmdOK(1).Enabled = False
'    '    cmdOK(2).Enabled = False
'    '    cmdOK(3).Enabled = False
'    '    ShowNoData
'    '    Screen.MousePointer = vbDefault
'    '    Me.Enabled = True
'    '    tmpBol = fnCancelNowFormAndShowParentForm(Me)
'    '    Exit Sub
'     If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
'        intCufaCnt = 0
'        adoRecordset.MoveFirst
'        Do While adoRecordset.EOF = False
'            '利益衝突案件：逐案號判斷
'            If PUB_ChkCufaByCase(Me.Name, Str02, "" & adoRecordset.Fields("本所案號"), "" & adoRecordset.Fields("cust01") & "," & adoRecordset.Fields("cust02") & "," & adoRecordset.Fields("cust03") & "," & adoRecordset.Fields("cust04") & "," & adoRecordset.Fields("cust05"), "" & adoRecordset.Fields("fcno")) = False Then
'                intCufaCnt = intCufaCnt + 1
'                adoRecordset.Delete
'            End If
'            adoRecordset.MoveNext
'        Loop
'        '利益衝突案件：限閱案件
'        If intCufaCnt > 0 Then
'            MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
'        End If
'        If adoRecordset.RecordCount = 0 Then
'              GoTo JumpToNoData
'        End If
'     End If
'     Set m_adoRst = adoRecordset.Clone
'    'end 2019/12/26
'Else
'JumpToNoData: 'Added by Lydia 2019/12/26
'   Set m_adoRst = adoRecordset.Clone 'Added by Lydia 2018/02/09
'   ShowNoData
'   Screen.MousePointer = vbDefault
'   Me.Enabled = True
'   tmpBol = fnCancelNowFormAndShowParentForm(Me)
'   Exit Sub
'End If
'
'grdDataList.FixedCols = 0
''Modified by Lydia 2018/12/17 中所反應耗時過久,卡在丟暫存檔(O8的寫法); O12 可以直接排序
''Set GrdDataList.Recordset = adoRecordset
'''Added by Lydia 2018/02/09 放到暫存檔,供Grid排序
''Set m_adoRst = PUB_CreateRecordset(adoRecordset, , , 300, Me.Name)
''Modified by Lydia 2018/12/22 拿掉desc
''m_adoRst.Sort = "FSort desc,本所案號 asc" 'Move by Lydia 2018/12/17 先排序,後指定資料集
'm_adoRst.Sort = "FSort ,本所案號 asc"
'SetRst2Grid
''end 2018/12/17
'm_blnColOrderAsc = True
''end 2018/02/09
'
'SetDataListWidth
'grdDataList.FixedCols = 4
'Me.Enabled = True
'End Sub

Private Sub Form_Unload(Cancel As Integer)
pub_QL05 = m_pub_QL05 'Add By Sindy 2025/9/12 還原此Form的查詢條件記錄 (多筆查詢有影響)
Set frm100102_2 = Nothing
End Sub

Private Sub grdDataList_SelChange()
grdDataList.Visible = False
grdDataList.row = grdDataList.MouseRow
grdDataList.col = 0
If grdDataList.row <> 0 Then
If grdDataList.Text = "V" Then
     grdDataList.Text = ""
     For i = 4 To grdDataList.Cols - 1
          grdDataList.col = i
          grdDataList.CellBackColor = QBColor(15)
    Next i
Else
     grdDataList.Text = "V"
     For i = 4 To grdDataList.Cols - 1
         grdDataList.col = i
         grdDataList.CellBackColor = &HFFC0C0
     Next i
End If
End If
grdDataList.Visible = True
End Sub

Sub StrMenu1() '關係案件
Dim Str01 As String, strTemp As Variant
Dim strArr(62) As String, StrOk(32) As String, StrOkTxt(12) As String
Dim strSQL2 As String
Dim StrSQL3 As String
Dim StrSQL4 As String
Dim strSQL5 As String
Dim StrSQL6 As String
Dim strSQL8 As String
Dim strField As String 'Add by Amy 2023/03/07
Dim dblRow As Double 'Add By Sindy 2025/9/3
   
   If Trim(Me.Tag) <> "" Then pub_QL05 = pub_QL05 & ";編號：" & Me.Tag & "(案件)" 'Add By Sindy 2025/8/13
   Me.Enabled = False
   If BolFrom100114 = False Then
      Str01 = ""    '申請人編號
      Str02 = ""    '系統類別
      Str03 = ""    '收文日期(起)
      Str04 = ""    '收文日期(迄)
      Str05 = ""    '案件性質(起)
      Str06 = ""    '案件性質(迄)
      Str07 = ""    '是否含來函資料
      Str01 = Me.Tag
      'Added by Lydia 2019/12/27 從對造案件->關係企業案，無法設定系統類別
      If m_Sys = "" And m_CaseNo <> "" Then
          Str02 = SystemNumber(m_CaseNo, 1)
      Else
      'end 2019/12/27
          Str02 = IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys))
      End If '2019/12/27
      If Trim(Str02) <> "" Then pub_QL05 = pub_QL05 & ";系統類別：" & Str02 'Add By Sindy 2025/8/13
      Str03 = m_Date1
      Str04 = m_Date2
      Str05 = m_Pty1
      Str06 = m_Pty2
      Str07 = m_CKind
      '組字串
      strSQL1 = ""
      
      '收文
      If m_Type = "1" Then
         If Len(Str03) <> 0 Then
             strSQL1 = strSQL1 + " and cp05>=" & Val(ChangeTStringToWString(Str03))
         End If
         If Len(Str04) <> 0 Then
             strSQL1 = strSQL1 + " and cp05<=" & Val(ChangeTStringToWString(Str04))
         End If
         'Add By Sindy 2025/8/13
         If Len(Str03) <> 0 Or Len(Str04) <> 0 Then
            pub_QL05 = pub_QL05 & ";收文日期：" & Str03 & "-" & Str04
         End If
         '2025/8/13 END
      'Add by Morgan 2008/11/26 配合代理人查詢畫面的條件
      '發文
      Else
         If Len(Str03) <> 0 Then
             strSQL1 = strSQL1 + " and cp27>=" & Val(ChangeTStringToWString(Str03))
         End If
         If Len(Str04) <> 0 Then
             strSQL1 = strSQL1 + " and cp27<=" & Val(ChangeTStringToWString(Str04))
         End If
         'Add By Sindy 2025/8/13
         If Len(Str03) <> 0 Or Len(Str04) <> 0 Then
            pub_QL05 = pub_QL05 & ";發文日期：" & Str03 & "-" & Str04
         End If
         '2025/8/13 END
      End If
      
      If Len(Str05) <> 0 Then
          strSQL1 = strSQL1 + " and cp10>='" & Str05 & "' "
      End If
      If Len(Str06) <> 0 Then
          strSQL1 = strSQL1 + " and cp10<='" & Str06 & "' "
      End If
      'Add By Sindy 2025/8/13
      If Len(Str05) <> 0 Or Len(Str06) <> 0 Then
         pub_QL05 = pub_QL05 & ";案件性質：" & Str05 & "-" & Str06
      End If
      '2025/8/13 END
      If UCase(Str07) = "N" Then
          strSQL1 = strSQL1 + " and cp09 < 'C' "
          pub_QL05 = pub_QL05 & ";是否含來函資料：不含" 'Add By Sindy 2025/8/13
      End If
   Else
         Str01 = ""    '申請人編號
         Str02 = ""    '系統類別
         Str01 = Me.Tag
         '組字串
         strSQL1 = ""
         strSQL2 = ""
         StrSQL3 = ""
         StrSQL4 = ""
         strSQL5 = ""
         StrSQL6 = ""
         If Len(Trim(m_Sys)) <> 0 Then
            strSQL1 = strSQL1 & " and tm01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 2) & ") "
            strSQL2 = strSQL2 & " and pa01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 1) & ") "
            StrSQL3 = StrSQL3 & " and sp01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 5) & ") "
            StrSQL4 = StrSQL4 & " and lc01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 3) & ") "
            strSQL8 = strSQL8 & " and hc01 in (" & SQLGrpStr(IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys)), 4) & ") "
            pub_QL05 = pub_QL05 & ";系統類別：" & Trim(m_Sys) 'Add By Sindy 2025/8/13
         End If
         If Len(Trim(m_Cty1)) <> 0 Then           '檢查申請國家
            strSQL1 = strSQL1 + " AND TM10='" & m_Cty1 & "' "
            strSQL2 = strSQL2 + " AND PA09='" & m_Cty1 & "' "
            StrSQL3 = StrSQL3 + " AND SP09='" & m_Cty1 & "' "
            StrSQL4 = StrSQL4 + " AND LC15='" & m_Cty1 & "' "
            pub_QL05 = pub_QL05 & ";申請國家：" & Trim(m_Cty1) 'Add By Sindy 2025/8/13
         End If
         If Len(Trim(m_Pty1)) <> 0 Then            '檢查案件性質
            strSQL1 = strSQL1 + " AND CP10>='" & m_Pty1 & "' "
            strSQL2 = strSQL2 + " AND CP10>='" & m_Pty1 & "' "
            StrSQL3 = StrSQL3 + " AND CP10>='" & m_Pty1 & "' "
            StrSQL4 = StrSQL4 + " AND CP10>='" & m_Pty1 & "' "
            strSQL8 = strSQL8 + " AND CP10>='" & m_Pty1 & "' "
         End If
         If Len(Trim(m_Pty2)) <> 0 Then
            strSQL1 = strSQL1 + " AND CP10<='" & m_Pty2 & "' "
            strSQL2 = strSQL2 + " AND CP10<='" & m_Pty2 & "' "
            StrSQL3 = StrSQL3 + " AND CP10<='" & m_Pty2 & "' "
            StrSQL4 = StrSQL4 + " AND CP10<='" & m_Pty2 & "' "
            strSQL8 = strSQL8 + " AND CP10<='" & m_Pty2 & "' "
         End If
         'Add By Sindy 2025/8/13
         If Len(m_Pty1) <> 0 Or Len(m_Pty2) <> 0 Then
            pub_QL05 = pub_QL05 & ";案件性質：" & m_Pty1 & "-" & m_Pty2
         End If
         '2025/8/13 END
         If m_Type = "1" Then        '收文
            If Len(m_Date1) <> 0 Then
               strSQL5 = strSQL5 + " AND CP05>=" & Val(ChangeTStringToWString(m_Date1)) & " "
            End If
            If Len(m_Date2) <> 0 Then
               strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(m_Date2)) & " "
            Else
               If Len(m_Date1) > 0 Then
                  strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(ServerDate - 19110000)) & " "
               End If
            End If
            'Add By Sindy 2025/8/13
            If Len(m_Date1) <> 0 Or Len(m_Date2) <> 0 Then
               pub_QL05 = pub_QL05 & ";收文日期：" & m_Date1 & "-" & m_Date2
            End If
            '2025/8/13 END
         Else
            If Len(m_Date1) <> 0 Then
               strSQL5 = strSQL5 + " AND CP27>=" & Val(ChangeTStringToWString(m_Date1)) & " "
            End If
            If Len(m_Date2) <> 0 Then
               strSQL5 = strSQL5 + " AND CP27<=" & Val(ChangeTStringToWString(m_Date2)) & " "
            Else
               If Len(m_Date1) > 0 Then
                  strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(ServerDate - 19110000)) & " "
               End If
            End If
            'Add By Sindy 2025/8/13
            If Len(m_Date1) <> 0 Or Len(m_Date2) <> 0 Then
               pub_QL05 = pub_QL05 & ";發文日期：" & m_Date1 & "-" & m_Date2
            End If
            '2025/8/13 END
         End If
   End If
    
CheckOC
'Add by Amy 2022/12/14 +e符號
'Modify by Amy 2023/03/06 bug-FCP-064401 該案無專用期且無領證發文,不應加e,避免未來資料多跑的慢優化語法
'T、FCT已有專用期間者,或無專用期但進度檔商標之註冊費717已發文
'strSQLE(0) = "Select tm01 as Ecp01,tm02 as Ecp02,tm03 as Ecp03,tm04 as Ecp04,'e' as EState From TradeMark " & _
'                    "Where tm136='1' And Tm04='00' And (tm21 is not null or (tm21 is null And Exists (Select * From CaseProgress Where cp01 in(" & SQLGrpStr2("", 2) & ") And cp10='717' And cp158<>0 And tm01=cp01(+) And tm02=cp02(+) And tm03=cp03(+) And tm04=cp04(+) )  ))"
'strSqlEW(0) = " And tm01=Ecp01(+) And tm02=Ecp02(+) And tm03=Ecp03(+) And tm04=Ecp04(+) "
strSQLE(0) = ""
strSqlEW(0) = Replace(Replace(專利進度註冊費已發文語法, "601", "717"), "pa", "tm")
strEField(0) = "Decode(tm136,'1',Decode(tm21,null,Decode(cp10,'717','e'),'e'))"
'P、FCP已有專用期間者,或無專用期但進度檔專利之領證601已發文
'strSQLE(1) = "Select pa01 as Ecp01,pa02 as Ecp02,pa03 as Ecp03,pa04 as Ecp04,'e' as EState From Patent " & _
'                    "Where pa178='1' And pa04='00' And (pa24 is not null or (pa24 is null And Exists (Select * From CaseProgress Where cp01 in(" & SQLGrpStr2("", 1) & ") And cp10='601' And cp158<>0 And pa01=cp01(+) And pa02=cp02(+) And pa03=cp03(+) And pa04=cp04(+) )  ))"
'strSqlEW(1) = " And pa01=Ecp01(+) And pa02=Ecp02(+) And pa03=Ecp03(+) And pa04=Ecp04(+) "
strSQLE(1) = ""
strSqlEW(1) = 專利進度註冊費已發文語法
strEField(1) = "Decode(pa178,'1',Decode(pa24,null,Decode(cp10,'601','e'),'e'))"
'end 2023/03/06
'end 2022/12/14
    
If bolIsL = False Then
'顯示申請人1
   grdDataList.ColWidth(9) = 1000
        'Added by Lydia 2019/12/26 利益衝突案件：於後面增加欄位
        SeColTM = " ,' ' as cnt,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno "
        SeColPA = " ,' ' as cnt,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno "
        SeColSP = " ,' ' as cnt,sp08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno "
        SeColLC = " ,' ' as cnt,lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno "
        SeColHC = " ,' ' as cnt,hc05 as cust01,hc24 as cust02,hc25 as cust03,hc26 as cust04,hc27 as cust05,'' as fcno "
        'end 2019/12/26
 
         'Modified by Morgan 2019/1/30 SQLGrpStr(Str02, #)->GetAddStr(str02) 取消不必要系統別的檢查(減少語法執行次數,分所有較大影響)
         'Modified by Lydia 2019/11/01 +增加欄位SeColTM
         'Modify by Amy 2022/12/14 +e符號 Nvl(EState,'')
         'Modify by Amy 2023/03/06 Nvl(EState,'')改為strEField(0),並優化e符號語法
         strSql = "SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','')||" & strEField(0) & " AS 本所案號 ,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,NVL(TM15,'') AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,TM09 AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度," & _
                  "NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort  " & SeColTM & _
                  " FROM TRADEMARK,nation,caseprogress,customer c1,customer c2,customer c3,customer c4,customer c5" & strSQLE(0) & " " & _
                  "WHERE tm10=na01(+) and TM23>='" & Mid(Me.Tag, 1, 6) & "000' AND TM23<='" & Mid(Me.Tag, 1, 6) & "zzz'  and tm01 in (" & GetAddStr(Str02) & ") and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and substR(tm23,1,8)=c1.cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=c1.cu02(+) " & _
               " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
               " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) " & strSqlEW(0) & strSQL1
         strSql = strSql & " union SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','')||" & strEField(0) & " AS 本所案號 ,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,NVL(TM15,'') AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,TM09 AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度," & _
                  "NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort  " & SeColTM & _
                  " FROM TRADEMARK,nation,caseprogress,customer c1,customer c2,customer c3,customer c4,customer c5" & strSQLE(0) & " " & _
                  "WHERE tm10=na01(+) and TM78>='" & Mid(Me.Tag, 1, 6) & "000' AND TM78<='" & Mid(Me.Tag, 1, 6) & "zzz'  and tm01 in (" & GetAddStr(Str02) & ") and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and substR(tm78,1,8)=c1.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c1.cu02(+) " & _
               " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
               " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) " & strSqlEW(0) & strSQL1
         strSql = strSql & " union SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','')||" & strEField(0) & " AS 本所案號 ,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,NVL(TM15,'') AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,TM09 AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度," & _
                  "NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort  " & SeColTM & _
                  " FROM TRADEMARK,nation,caseprogress,customer c1,customer c2,customer c3,customer c4,customer c5" & strSQLE(0) & " " & _
                  "WHERE tm10=na01(+) and TM79>='" & Mid(Me.Tag, 1, 6) & "000' AND TM79<='" & Mid(Me.Tag, 1, 6) & "zzz'  and tm01 in (" & GetAddStr(Str02) & ") and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and substR(tm79,1,8)=c1.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c1.cu02(+) " & _
               " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
               " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) " & strSqlEW(0) & strSQL1
         strSql = strSql & " union SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','')||" & strEField(0) & " AS 本所案號 ,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,NVL(TM15,'') AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,TM09 AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度," & _
                  "NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort  " & SeColTM & _
                  " FROM TRADEMARK,nation,caseprogress,customer c1,customer c2,customer c3,customer c4,customer c5" & strSQLE(0) & " " & _
                  "WHERE tm10=na01(+) and TM80>='" & Mid(Me.Tag, 1, 6) & "000' AND TM80<='" & Mid(Me.Tag, 1, 6) & "zzz'  and tm01 in (" & GetAddStr(Str02) & ") and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and substR(tm80,1,8)=c1.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c1.cu02(+) " & _
               " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
               " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) " & strSqlEW(0) & strSQL1
         strSql = strSql & " union SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','')||" & strEField(0) & " AS 本所案號 ,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,NVL(TM15,'') AS 審定專利號數,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,TM09 AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度," & _
                  "NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort  " & SeColTM & _
                  " FROM TRADEMARK,nation,caseprogress,customer c1,customer c2,customer c3,customer c4,customer c5" & strSQLE(0) & " " & _
                  "WHERE tm10=na01(+) and TM81>='" & Mid(Me.Tag, 1, 6) & "000' AND TM81<='" & Mid(Me.Tag, 1, 6) & "zzz'  and tm01 in (" & GetAddStr(Str02) & ") and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and substR(tm81,1,8)=c1.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c1.cu02(+) " & _
               " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+) " & _
               " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+) and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+) and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+) " & strSqlEW(0) & strSQL1
         
         'Modify By Sindy 2014/7/7 +||DECODE(PA165,'Y','＃','')
         'Modified by Lydia 2019/11/01 +增加欄位SeColPA
         'Modify by Amy 2022/12/14 +e符號 Nvl(EState,'')
         'Modify by Amy 2023/03/06 Nvl(EState,'')改為strEField(1),並優化e符號語法
         strSql = strSql + "union select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','')||" & strEField(1) & " AS 本所案號 ,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,NVL(PA22,'') AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
                  "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度,NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2,NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3,NVL(c4.CU04,NVL(c4.CU05||c4.CU88||c4.CU89||c4.CU90,c4.CU06)) AS 申請人4,NVL(c5.CU04,NVL(c5.CU05||c5.CU88||c5.CU89||c5.CU90,c5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort " & SeColPA & _
                  " FROM PATENT,nation,caseprogress,customer c1,customer c2,customer c3,customer c4,customer c5" & strSQLE(1) & " " & _
                  " WHERE pa09=na01(+) and PA26>='" & Mid(Me.Tag, 1, 6) & "000' AND PA26<='" & Mid(Me.Tag, 1, 6) & "zzz' and pa01 in (" & GetAddStr(Str02) & ") and substr(pa26,1,8)=c1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=c1.cu02(+) and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substR(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) " & strSqlEW(1) & strSQL1
         strSql = strSql + "union select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','')||" & strEField(1) & " AS 本所案號 ,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,NVL(PA22,'') AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
                  "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度,NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2,NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3,NVL(c4.CU04,NVL(c4.CU05||c4.CU88||c4.CU89||c4.CU90,c4.CU06)) AS 申請人4,NVL(c5.CU04,NVL(c5.CU05||c5.CU88||c5.CU89||c5.CU90,c5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort " & SeColPA & _
                  " FROM PATENT,nation,caseprogress,customer c1,customer c2,customer c3,customer c4,customer c5" & strSQLE(1) & " " & _
                  " WHERE pa09=na01(+) and PA27>='" & Mid(Me.Tag, 1, 6) & "000' AND PA27<='" & Mid(Me.Tag, 1, 6) & "zzz' and pa01 in (" & GetAddStr(Str02) & ") and substr(pa26,1,8)=c1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=c1.cu02(+) and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substR(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) " & strSqlEW(1) & strSQL1
         strSql = strSql + "union select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','')||" & strEField(1) & " AS 本所案號 ,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,NVL(PA22,'') AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
                  "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度,NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2,NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3,NVL(c4.CU04,NVL(c4.CU05||c4.CU88||c4.CU89||c4.CU90,c4.CU06)) AS 申請人4,NVL(c5.CU04,NVL(c5.CU05||c5.CU88||c5.CU89||c5.CU90,c5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort " & SeColPA & _
                  " FROM PATENT,nation,caseprogress,customer c1,customer c2,customer c3,customer c4,customer c5" & strSQLE(1) & " " & _
                  " WHERE pa09=na01(+) and PA28>='" & Mid(Me.Tag, 1, 6) & "000' AND PA28<='" & Mid(Me.Tag, 1, 6) & "zzz' and pa01 in (" & GetAddStr(Str02) & ") and substr(pa26,1,8)=c1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=c1.cu02(+) and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substR(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) " & strSqlEW(1) & strSQL1
         strSql = strSql + "union select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','')||" & strEField(1) & " AS 本所案號 ,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,NVL(PA22,'') AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
                  "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度,NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2,NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3,NVL(c4.CU04,NVL(c4.CU05||c4.CU88||c4.CU89||c4.CU90,c4.CU06)) AS 申請人4,NVL(c5.CU04,NVL(c5.CU05||c5.CU88||c5.CU89||c5.CU90,c5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort " & SeColPA & _
                  " FROM PATENT,nation,caseprogress,customer c1,customer c2,customer c3,customer c4,customer c5" & strSQLE(1) & " " & _
                  " WHERE pa09=na01(+) and PA29>='" & Mid(Me.Tag, 1, 6) & "000' AND PA29<='" & Mid(Me.Tag, 1, 6) & "zzz' and pa01 in (" & GetAddStr(Str02) & ") and substr(pa26,1,8)=c1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=c1.cu02(+) and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substR(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) " & strSqlEW(1) & strSQL1
         strSql = strSql + "union select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','')||" & strEField(1) & " AS 本所案號 ,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,NVL(PA22,'') AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
                  "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度,NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2,NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3,NVL(c4.CU04,NVL(c4.CU05||c4.CU88||c4.CU89||c4.CU90,c4.CU06)) AS 申請人4,NVL(c5.CU04,NVL(c5.CU05||c5.CU88||c5.CU89||c5.CU90,c5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort " & SeColPA & _
                  " FROM PATENT,nation,caseprogress,customer c1,customer c2,customer c3,customer c4,customer c5" & strSQLE(1) & " " & _
                  " WHERE pa09=na01(+) and PA30>='" & Mid(Me.Tag, 1, 6) & "000' AND PA30<='" & Mid(Me.Tag, 1, 6) & "zzz' and pa01 in (" & GetAddStr(Str02) & ") and substr(pa26,1,8)=c1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=c1.cu02(+) and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) and substR(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) " & strSqlEW(1) & strSQL1

         'Modified by Lydia 2019/11/01 +增加欄位SeColSP
         'Modify by Amy 2020/02/05 +SP73 商品類別
         strSql = strSql + "union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
                  "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
                  ",NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2,NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort " & SeColSP & _
                  " FROM SERVICEPRACTICE,nation,caseprogress,customer c1,customer c2,customer c3,customer c4,customer c5 WHERE sp09=na01(+) and SP08>='" & Mid(Me.Tag, 1, 6) & "000' AND SP08<='" & Mid(Me.Tag, 1, 6) & "zzz' and sp01 in (" & GetAddStr(Str02) & ") and substr(sp08,1,8)=c1.cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=c1.cu02(+) and substr(sp58,1,8)=c2.cu01(+) and decode(substr(sp58,9,1),null,'0',substr(sp58,9,1))=c2.cu02(+) and substr(sp59,1,8)=c3.cu01(+) and decode(substr(sp59,9,1),null,'0',substr(sp59,9,1))=c3.cu02(+) " & _
                  " and substr(sp65,1,8)=c4.cu01(+) and decode(substr(sp65,9,1),null,'0',substr(sp65,9,1))=c4.cu02(+) and substr(sp66,1,8)=c5.cu01(+) and decode(substr(sp66,9,1),null,'0',substr(sp66,9,1))=c5.cu02(+) and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) " & strSQL1
         strSql = strSql + "union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
                  "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
                  ",NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2,NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort " & SeColSP & _
                  " FROM SERVICEPRACTICE,nation,caseprogress,customer c1,customer c2,customer c3,customer c4,customer c5 WHERE sp09=na01(+) and SP58>='" & Mid(Me.Tag, 1, 6) & "000' AND SP58<='" & Mid(Me.Tag, 1, 6) & "zzz' and sp01 in (" & GetAddStr(Str02) & ") and substr(sp08,1,8)=c1.cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=c1.cu02(+) and substr(sp58,1,8)=c2.cu01(+) and decode(substr(sp58,9,1),null,'0',substr(sp58,9,1))=c2.cu02(+) and substr(sp59,1,8)=c3.cu01(+) and decode(substr(sp59,9,1),null,'0',substr(sp59,9,1))=c3.cu02(+) " & _
                  " and substr(sp65,1,8)=c4.cu01(+) and decode(substr(sp65,9,1),null,'0',substr(sp65,9,1))=c4.cu02(+) and substr(sp66,1,8)=c5.cu01(+) and decode(substr(sp66,9,1),null,'0',substr(sp66,9,1))=c5.cu02(+) and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) " & strSQL1
         strSql = strSql + "union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
                  "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
                  ",NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2,NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort " & SeColSP & _
                  " FROM SERVICEPRACTICE,nation,caseprogress,customer c1,customer c2,customer c3,customer c4,customer c5 WHERE sp09=na01(+) and SP59>='" & Mid(Me.Tag, 1, 6) & "000' AND SP59<='" & Mid(Me.Tag, 1, 6) & "zzz' and sp01 in (" & GetAddStr(Str02) & ") and substr(sp08,1,8)=c1.cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=c1.cu02(+) and substr(sp58,1,8)=c2.cu01(+) and decode(substr(sp58,9,1),null,'0',substr(sp58,9,1))=c2.cu02(+) and substr(sp59,1,8)=c3.cu01(+) and decode(substr(sp59,9,1),null,'0',substr(sp59,9,1))=c3.cu02(+) " & _
                  " and substr(sp65,1,8)=c4.cu01(+) and decode(substr(sp65,9,1),null,'0',substr(sp65,9,1))=c4.cu02(+) and substr(sp66,1,8)=c5.cu01(+) and decode(substr(sp66,9,1),null,'0',substr(sp66,9,1))=c5.cu02(+) and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) " & strSQL1
         strSql = strSql + "union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
                  "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
                  ",NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2,NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort " & SeColSP & _
                  " FROM SERVICEPRACTICE,nation,caseprogress,customer c1,customer c2,customer c3,customer c4,customer c5 WHERE sp09=na01(+) and SP65>='" & Mid(Me.Tag, 1, 6) & "000' AND SP65<='" & Mid(Me.Tag, 1, 6) & "zzz' and sp01 in (" & GetAddStr(Str02) & ") and substr(sp08,1,8)=c1.cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=c1.cu02(+) and substr(sp58,1,8)=c2.cu01(+) and decode(substr(sp58,9,1),null,'0',substr(sp58,9,1))=c2.cu02(+) and substr(sp59,1,8)=c3.cu01(+) and decode(substr(sp59,9,1),null,'0',substr(sp59,9,1))=c3.cu02(+) " & _
                  " and substr(sp65,1,8)=c4.cu01(+) and decode(substr(sp65,9,1),null,'0',substr(sp65,9,1))=c4.cu02(+) and substr(sp66,1,8)=c5.cu01(+) and decode(substr(sp66,9,1),null,'0',substr(sp66,9,1))=c5.cu02(+) and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) " & strSQL1
         strSql = strSql + "union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
                  "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
                  ",NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2,NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort " & SeColSP & _
                  " FROM SERVICEPRACTICE,nation,caseprogress,customer c1,customer c2,customer c3,customer c4,customer c5 WHERE sp09=na01(+) and SP66>='" & Mid(Me.Tag, 1, 6) & "000' AND SP66<='" & Mid(Me.Tag, 1, 6) & "zzz' and sp01 in (" & GetAddStr(Str02) & ") and substr(sp08,1,8)=c1.cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=c1.cu02(+) and substr(sp58,1,8)=c2.cu01(+) and decode(substr(sp58,9,1),null,'0',substr(sp58,9,1))=c2.cu02(+) and substr(sp59,1,8)=c3.cu01(+) and decode(substr(sp59,9,1),null,'0',substr(sp59,9,1))=c3.cu02(+) " & _
                  " and substr(sp65,1,8)=c4.cu01(+) and decode(substr(sp65,9,1),null,'0',substr(sp65,9,1))=c4.cu02(+) and substr(sp66,1,8)=c5.cu01(+) and decode(substr(sp66,9,1),null,'0',substr(sp66,9,1))=c5.cu02(+) and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) " & strSQL1
         'end 2020/02/05
         
         'Modify By Sindy 2011/1/20 +LC43,LC44,LC45,LC46
         'Modify by Amy 2018/08/15 +CasePropertyMap 若專案服務案 lc52=Y 顯示 案件進度+案件性質
         'Modified by Lydia 2019/11/01 +增加欄位SeColLC
         strSql = strSql + "union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,Decode(LC52,'Y',CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')',NVL(LC05,NVL(LC06,LC07)))AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,'' AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,'' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2,NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
                  " FROM LAWCASE,nation,caseprogress,customer c1,customer c2,customer c3,customer c4,customer c5,CasePropertyMap " & _
                  " WHERE lc15=na01(+) and LC11>='" & Mid(Me.Tag, 1, 6) & "000' and LC11<='" & Mid(Me.Tag, 1, 6) & "zzz' and lc01 in (" & GetAddStr(Str02) & ") " & _
                  " and substr(lc11,1,8)=c1.cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=c1.cu02(+) " & _
                  " and substr(lc43,1,8)=c2.cu01(+) and decode(substr(lc43,9,1),null,'0',substr(lc43,9,1))=c2.cu02(+) " & _
                  " and substr(lc44,1,8)=c3.cu01(+) and decode(substr(lc44,9,1),null,'0',substr(lc44,9,1))=c3.cu02(+) " & _
                  " and substr(lc45,1,8)=c4.cu01(+) and decode(substr(lc45,9,1),null,'0',substr(lc45,9,1))=c4.cu02(+) " & _
                  " and substr(lc46,1,8)=c5.cu01(+) and decode(substr(lc46,9,1),null,'0',substr(lc46,9,1))=c5.cu02(+) " & _
                  " and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1
         strSql = strSql + "union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,Decode(LC52,'Y',CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')',NVL(LC05,NVL(LC06,LC07)))AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,'' AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,'' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2,NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
                  " FROM LAWCASE,nation,caseprogress,customer c1,customer c2,customer c3,customer c4,customer c5,CasePropertyMap " & _
                  " WHERE lc15=na01(+) and LC43>='" & Mid(Me.Tag, 1, 6) & "000' and LC43<='" & Mid(Me.Tag, 1, 6) & "zzz' and lc01 in (" & GetAddStr(Str02) & ") " & _
                  " and substr(lc11,1,8)=c1.cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=c1.cu02(+) " & _
                  " and substr(lc43,1,8)=c2.cu01(+) and decode(substr(lc43,9,1),null,'0',substr(lc43,9,1))=c2.cu02(+) " & _
                  " and substr(lc44,1,8)=c3.cu01(+) and decode(substr(lc44,9,1),null,'0',substr(lc44,9,1))=c3.cu02(+) " & _
                  " and substr(lc45,1,8)=c4.cu01(+) and decode(substr(lc45,9,1),null,'0',substr(lc45,9,1))=c4.cu02(+) " & _
                  " and substr(lc46,1,8)=c5.cu01(+) and decode(substr(lc46,9,1),null,'0',substr(lc46,9,1))=c5.cu02(+) " & _
                  " and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1
         strSql = strSql + "union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,Decode(LC52,'Y',CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')',NVL(LC05,NVL(LC06,LC07)))AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,'' AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,'' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2,NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
                  " FROM LAWCASE,nation,caseprogress,customer c1,customer c2,customer c3,customer c4,customer c5,CasePropertyMap " & _
                  " WHERE lc15=na01(+) and LC44>='" & Mid(Me.Tag, 1, 6) & "000' and LC44<='" & Mid(Me.Tag, 1, 6) & "zzz' and lc01 in (" & GetAddStr(Str02) & ") " & _
                  " and substr(lc11,1,8)=c1.cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=c1.cu02(+) " & _
                  " and substr(lc43,1,8)=c2.cu01(+) and decode(substr(lc43,9,1),null,'0',substr(lc43,9,1))=c2.cu02(+) " & _
                  " and substr(lc44,1,8)=c3.cu01(+) and decode(substr(lc44,9,1),null,'0',substr(lc44,9,1))=c3.cu02(+) " & _
                  " and substr(lc45,1,8)=c4.cu01(+) and decode(substr(lc45,9,1),null,'0',substr(lc45,9,1))=c4.cu02(+) " & _
                  " and substr(lc46,1,8)=c5.cu01(+) and decode(substr(lc46,9,1),null,'0',substr(lc46,9,1))=c5.cu02(+) " & _
                  " and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1
         strSql = strSql + "union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,Decode(LC52,'Y',CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')',NVL(LC05,NVL(LC06,LC07)))AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,'' AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,'' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2,NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
                  " FROM LAWCASE,nation,caseprogress,customer c1,customer c2,customer c3,customer c4,customer c5,CasePropertyMap " & _
                  " WHERE lc15=na01(+) and LC45>='" & Mid(Me.Tag, 1, 6) & "000' and LC45<='" & Mid(Me.Tag, 1, 6) & "zzz' and lc01 in (" & GetAddStr(Str02) & ") " & _
                  " and substr(lc11,1,8)=c1.cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=c1.cu02(+) " & _
                  " and substr(lc43,1,8)=c2.cu01(+) and decode(substr(lc43,9,1),null,'0',substr(lc43,9,1))=c2.cu02(+) " & _
                  " and substr(lc44,1,8)=c3.cu01(+) and decode(substr(lc44,9,1),null,'0',substr(lc44,9,1))=c3.cu02(+) " & _
                  " and substr(lc45,1,8)=c4.cu01(+) and decode(substr(lc45,9,1),null,'0',substr(lc45,9,1))=c4.cu02(+) " & _
                  " and substr(lc46,1,8)=c5.cu01(+) and decode(substr(lc46,9,1),null,'0',substr(lc46,9,1))=c5.cu02(+) " & _
                  " and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1
         strSql = strSql + "union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,Decode(LC52,'Y',CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')',NVL(LC05,NVL(LC06,LC07)))AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,'' AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,'' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2,NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
                  " FROM LAWCASE,nation,caseprogress,customer c1,customer c2,customer c3,customer c4,customer c5,CasePropertyMap " & _
                  " WHERE lc15=na01(+) and LC46>='" & Mid(Me.Tag, 1, 6) & "000' and LC46<='" & Mid(Me.Tag, 1, 6) & "zzz' and lc01 in (" & GetAddStr(Str02) & ") " & _
                  " and substr(lc11,1,8)=c1.cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=c1.cu02(+) " & _
                  " and substr(lc43,1,8)=c2.cu01(+) and decode(substr(lc43,9,1),null,'0',substr(lc43,9,1))=c2.cu02(+) " & _
                  " and substr(lc44,1,8)=c3.cu01(+) and decode(substr(lc44,9,1),null,'0',substr(lc44,9,1))=c3.cu02(+) " & _
                  " and substr(lc45,1,8)=c4.cu01(+) and decode(substr(lc45,9,1),null,'0',substr(lc45,9,1))=c4.cu02(+) " & _
                  " and substr(lc46,1,8)=c5.cu01(+) and decode(substr(lc46,9,1),null,'0',substr(lc46,9,1))=c5.cu02(+) " & _
                  " and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & strSQL1
                  
         'Modify By Sindy 2011/1/20 +HC24,HC25,HC26,HC27
         'Modified by Lydia 2019/11/01 +增加欄位SeColHC
         strSql = strSql + "union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,' ' AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,'' AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,'' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2,NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') as FSort " & SeColHC & _
                  " FROM HIRECASE,caseprogress,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                  " WHERE HC05>='" & Mid(Me.Tag, 1, 6) & "000' AND HC05<='" & Mid(Me.Tag, 1, 6) & "zzz' and hc01 in (" & GetAddStr(Str02) & ") " & _
                  " and substr(hc05,1,8)=c1.cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=c1.cu02(+) " & _
                  " and substr(hc24,1,8)=c2.cu01(+) and decode(substr(hc24,9,1),null,'0',substr(hc24,9,1))=c2.cu02(+) " & _
                  " and substr(hc25,1,8)=c3.cu01(+) and decode(substr(hc25,9,1),null,'0',substr(hc25,9,1))=c3.cu02(+) " & _
                  " and substr(hc26,1,8)=c4.cu01(+) and decode(substr(hc26,9,1),null,'0',substr(hc26,9,1))=c4.cu02(+) " & _
                  " and substr(hc27,1,8)=c5.cu01(+) and decode(substr(hc27,9,1),null,'0',substr(hc27,9,1))=c5.cu02(+) " & _
                  " and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) " & strSQL1
         strSql = strSql + "union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,' ' AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,'' AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,'' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2,NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') as FSort " & SeColHC & _
                  " FROM HIRECASE,caseprogress,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                  " WHERE HC24>='" & Mid(Me.Tag, 1, 6) & "000' AND HC24<='" & Mid(Me.Tag, 1, 6) & "zzz' and hc01 in (" & GetAddStr(Str02) & ") " & _
                  " and substr(hc05,1,8)=c1.cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=c1.cu02(+) " & _
                  " and substr(hc24,1,8)=c2.cu01(+) and decode(substr(hc24,9,1),null,'0',substr(hc24,9,1))=c2.cu02(+) " & _
                  " and substr(hc25,1,8)=c3.cu01(+) and decode(substr(hc25,9,1),null,'0',substr(hc25,9,1))=c3.cu02(+) " & _
                  " and substr(hc26,1,8)=c4.cu01(+) and decode(substr(hc26,9,1),null,'0',substr(hc26,9,1))=c4.cu02(+) " & _
                  " and substr(hc27,1,8)=c5.cu01(+) and decode(substr(hc27,9,1),null,'0',substr(hc27,9,1))=c5.cu02(+) " & _
                  " and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) " & strSQL1
         strSql = strSql + "union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,' ' AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,'' AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,'' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2,NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') as FSort " & SeColHC & _
                  " FROM HIRECASE,caseprogress,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                  " WHERE HC25>='" & Mid(Me.Tag, 1, 6) & "000' AND HC25<='" & Mid(Me.Tag, 1, 6) & "zzz' and hc01 in (" & GetAddStr(Str02) & ") " & _
                  " and substr(hc05,1,8)=c1.cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=c1.cu02(+) " & _
                  " and substr(hc24,1,8)=c2.cu01(+) and decode(substr(hc24,9,1),null,'0',substr(hc24,9,1))=c2.cu02(+) " & _
                  " and substr(hc25,1,8)=c3.cu01(+) and decode(substr(hc25,9,1),null,'0',substr(hc25,9,1))=c3.cu02(+) " & _
                  " and substr(hc26,1,8)=c4.cu01(+) and decode(substr(hc26,9,1),null,'0',substr(hc26,9,1))=c4.cu02(+) " & _
                  " and substr(hc27,1,8)=c5.cu01(+) and decode(substr(hc27,9,1),null,'0',substr(hc27,9,1))=c5.cu02(+) " & _
                  " and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) " & strSQL1
         strSql = strSql + "union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,' ' AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,'' AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,'' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2,NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') as FSort " & SeColHC & _
                  " FROM HIRECASE,caseprogress,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                  " WHERE HC26>='" & Mid(Me.Tag, 1, 6) & "000' AND HC26<='" & Mid(Me.Tag, 1, 6) & "zzz' and hc01 in (" & GetAddStr(Str02) & ") " & _
                  " and substr(hc05,1,8)=c1.cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=c1.cu02(+) " & _
                  " and substr(hc24,1,8)=c2.cu01(+) and decode(substr(hc24,9,1),null,'0',substr(hc24,9,1))=c2.cu02(+) " & _
                  " and substr(hc25,1,8)=c3.cu01(+) and decode(substr(hc25,9,1),null,'0',substr(hc25,9,1))=c3.cu02(+) " & _
                  " and substr(hc26,1,8)=c4.cu01(+) and decode(substr(hc26,9,1),null,'0',substr(hc26,9,1))=c4.cu02(+) " & _
                  " and substr(hc27,1,8)=c5.cu01(+) and decode(substr(hc27,9,1),null,'0',substr(hc27,9,1))=c5.cu02(+) " & _
                  " and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) " & strSQL1
         strSql = strSql + "union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,' ' AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,'' AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,'' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,NVL(c2.CU04,NVL(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2,NVL(c3.CU04,NVL(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') as FSort " & SeColHC & _
                  " FROM HIRECASE,caseprogress,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                  " WHERE HC27>='" & Mid(Me.Tag, 1, 6) & "000' AND HC27<='" & Mid(Me.Tag, 1, 6) & "zzz' and hc01 in (" & GetAddStr(Str02) & ") " & _
                  " and substr(hc05,1,8)=c1.cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=c1.cu02(+) " & _
                  " and substr(hc24,1,8)=c2.cu01(+) and decode(substr(hc24,9,1),null,'0',substr(hc24,9,1))=c2.cu02(+) " & _
                  " and substr(hc25,1,8)=c3.cu01(+) and decode(substr(hc25,9,1),null,'0',substr(hc25,9,1))=c3.cu02(+) " & _
                  " and substr(hc26,1,8)=c4.cu01(+) and decode(substr(hc26,9,1),null,'0',substr(hc26,9,1))=c4.cu02(+) " & _
                  " and substr(hc27,1,8)=c5.cu01(+) and decode(substr(hc27,9,1),null,'0',substr(hc27,9,1))=c5.cu02(+) " & _
                  " and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) " & strSQL1
Else   '法務進度
        'Added by Lydia 2019/12/26 利益衝突案件：於後面增加欄位
        SeColTM = " ,' ' as CP09,' ' as cnt,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno "
        SeColPA = " ,' ' as CP09,' ' as cnt,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno "
        SeColSP = " ,' ' as CP09,' ' as cnt,sp08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno "
        SeColLC = " ,' ' as CP09,' ' as cnt,lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno "
        SeColHC = " ,' ' as CP09,' ' as cnt,hc05 as cust01,hc24 as cust02,hc25 as cust03,hc26 as cust04,hc27 as cust05,'' as fcno "
        'end 2019/12/26
        
         'Modify By Sindy 2011/1/21 +LC43,LC44,LC45,LC46
         'Modify by Amy 2018/08/15 若專案服務案 lc52=Y 顯示 案件進度+案件性質
         'Modified by Lydia 2019/11/01 +增加欄位SeColLC
         strSql = "select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,Decode(LC52,'Y',CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')',NVL(LC05,NVL(LC06,LC07)))AS 案件名稱,na03 AS 申請國家, ' ' AS 申請日 ,' ' AS 准駁 ,   sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
                  " FROM LAWCASE,nation,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
                  " WHERE lc15=na01(+) " & _
                  " and LC11>='" & Mid(Me.Tag, 1, 6) & "000' and LC11<='" & Mid(Me.Tag, 1, 6) & "zzz' and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00' " & _
                  " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1)) = CU02(+) " & _
                  " and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+) " & _
                  " and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & StrSQL4 & strSQL5
         strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,Decode(LC52,'Y',CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')',NVL(LC05,NVL(LC06,LC07)))AS 案件名稱,na03 AS 申請國家, ' ' AS 申請日 ,' ' AS 准駁 ,   sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
                  " FROM LAWCASE,nation,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
                  " WHERE lc15=na01(+) " & _
                  " and LC43>='" & Mid(Me.Tag, 1, 6) & "000' and LC43<='" & Mid(Me.Tag, 1, 6) & "zzz' and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00' " & _
                  " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1)) = CU02(+) " & _
                  " and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+) " & _
                  " and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & StrSQL4 & strSQL5
         strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,Decode(LC52,'Y',CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')',NVL(LC05,NVL(LC06,LC07)))AS 案件名稱,na03 AS 申請國家, ' ' AS 申請日 ,' ' AS 准駁 ,   sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
                  " FROM LAWCASE,nation,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
                  " WHERE lc15=na01(+) " & _
                  " and LC44>='" & Mid(Me.Tag, 1, 6) & "000' and LC44<='" & Mid(Me.Tag, 1, 6) & "zzz' and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00' " & _
                  " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1)) = CU02(+) " & _
                  " and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+) " & _
                  " and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & StrSQL4 & strSQL5
         strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,Decode(LC52,'Y',CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')',NVL(LC05,NVL(LC06,LC07)))AS 案件名稱,na03 AS 申請國家, ' ' AS 申請日 ,' ' AS 准駁 ,   sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
                  " FROM LAWCASE,nation,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
                  " WHERE lc15=na01(+) " & _
                  " and LC45>='" & Mid(Me.Tag, 1, 6) & "000' and LC45<='" & Mid(Me.Tag, 1, 6) & "zzz' and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00' " & _
                  " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1)) = CU02(+) " & _
                  " and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+) " & _
                  " and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & StrSQL4 & strSQL5
         strSql = strSql + " union select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號 , DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號 ,Decode(LC52,'Y',CP64||'('||NVL(DECODE(LC15,'000',CPM03,CPM04),CP10)||')',NVL(LC05,NVL(LC06,LC07)))AS 案件名稱,na03 AS 申請國家, ' ' AS 申請日 ,' ' AS 准駁 ,   sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') as FSort " & SeColLC & _
                  " FROM LAWCASE,nation,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
                  " WHERE lc15=na01(+) " & _
                  " and LC46>='" & Mid(Me.Tag, 1, 6) & "000' and LC46<='" & Mid(Me.Tag, 1, 6) & "zzz' and lc01 in (" & GetAddStr(Str02) & ") AND LC04='00' " & _
                  " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1)) = CU02(+) " & _
                  " and lC01=Cp01(+) AND lC02=Cp02(+) AND lC03=Cp03(+) AND lC04=Cp04(+) " & _
                  " and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & StrSQL4 & strSQL5
         
         'Modify By Sindy 2011/1/21 +HC24,HC25,HC26,HC27
         'Modified by Lydia 2019/11/01 +增加欄位SeColHC
         strSql = strSql + " union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,decode(cp10,'0',sqldatet(cp53),'台灣') AS 申請國家, decode(cp10,'0',sqldatet(cp54),'') AS 申請日 ,' ' AS 准駁 ,   sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort " & SeColHC & _
                  " FROM HIRECASE,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
                  " WHERE HC05>='" & Mid(Me.Tag, 1, 6) & "000' AND HC05<='" & Mid(Me.Tag, 1, 6) & "zzz' and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00' " & _
                  " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+) " & _
                  " and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) " & _
                  " and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL8 & strSQL5
         strSql = strSql + " union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,decode(cp10,'0',sqldatet(cp53),'台灣') AS 申請國家, decode(cp10,'0',sqldatet(cp54),'') AS 申請日 ,' ' AS 准駁 ,   sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort " & SeColHC & _
                  " FROM HIRECASE,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
                  " WHERE HC24>='" & Mid(Me.Tag, 1, 6) & "000' AND HC24<='" & Mid(Me.Tag, 1, 6) & "zzz' and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00' " & _
                  " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+) " & _
                  " and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) " & _
                  " and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL8 & strSQL5
         strSql = strSql + " union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,decode(cp10,'0',sqldatet(cp53),'台灣') AS 申請國家, decode(cp10,'0',sqldatet(cp54),'') AS 申請日 ,' ' AS 准駁 ,   sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort " & SeColHC & _
                  " FROM HIRECASE,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
                  " WHERE HC25>='" & Mid(Me.Tag, 1, 6) & "000' AND HC25<='" & Mid(Me.Tag, 1, 6) & "zzz' and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00' " & _
                  " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+) " & _
                  " and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) " & _
                  " and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL8 & strSQL5
         strSql = strSql + " union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,decode(cp10,'0',sqldatet(cp53),'台灣') AS 申請國家, decode(cp10,'0',sqldatet(cp54),'') AS 申請日 ,' ' AS 准駁 ,   sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort " & SeColHC & _
                  " FROM HIRECASE,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
                  " WHERE HC26>='" & Mid(Me.Tag, 1, 6) & "000' AND HC26<='" & Mid(Me.Tag, 1, 6) & "zzz' and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00' " & _
                  " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+) " & _
                  " and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) " & _
                  " and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL8 & strSQL5
         strSql = strSql + " union select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,decode(cp10,'0',sqldatet(cp53),'台灣') AS 申請國家, decode(cp10,'0',sqldatet(cp54),'') AS 申請日 ,' ' AS 准駁 ,   sqldatet(cp05) AS 收文日 ,nvl(cpm03,cpm04) AS 案件性質,s1.st02 as 智權人員,s2.st02 as 承辦人,sqldatet(cp06) as 本所期限,sqldatet(cp07) as 法定期限,sqldatet(cp27) as 發文日,sqldatet(cp57) as 取消收文日," & cntFaSql & " as 代理人,decode(cp24,'1','准','2','駁',' ') AS 結果,NVL(CP40,NVL(CP50,DECODE(CP56,CU01||CU02,CU04))) as 相關人,cp64 AS 進度備註,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort " & SeColHC & _
                  " FROM HIRECASE,CUSTOMER,caseprogress,casepropertymap,staff s1,staff s2,fagent,SystemKind " & _
                  " WHERE HC27>='" & Mid(Me.Tag, 1, 6) & "000' AND HC27<='" & Mid(Me.Tag, 1, 6) & "zzz' and hc01 in (" & GetAddStr(Str02) & ") AND HC04='00' " & _
                  " AND SUBSTR(cp56,1,8)=CU01(+) AND DECODE(SUBSTR(cp56,9,1),NULL,'0',SUBSTR(cp56,9,1))=CU02(+) " & _
                  " and hc01=cp01(+) and hc02=cp02(+) and hc03=cp03(+) and hc04=cp04(+) " & _
                  " and cp13=s1.st01(+) and cp14=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP01=SK01(+) " & strSQL8 & strSQL5
End If

'Modify by Amy 2023/03/07 勾PCT,'' as 審定專利號數 未取代會錯
strField = ",審定專利號數"
If ChkPCT.Value = vbChecked Then
    strField = ",PCT"
    strSql = Replace(Replace(Replace(Replace(Replace(UCase(strSql), "NVL(TM15,'') AS 審定專利號數", "'' as PCT"), "NVL(SP14,SP13) AS 審定專利號數", "'' as PCT"), "'' AS 審定專利號數", "'' as PCT"), "NVL(PA22,'') AS 審定專利號數", "pa46 as PCT"), "'' AS 審定專利號數", "'' as PCT")
End If

'Modify by Amy 2020/06/16 案件名稱顯示40字,避免排序「案件名稱」會Error ex:申請人查X43988060->案件 鈕->關係企業案 鈕
strSql = "Select V,本所案號,分所號,SubStr(案件名稱,1,40) as 案件名稱,申請國家,申請案號,申請日" & strField & ",准駁,申請人1,商品類別,專用期間,專利公告號,最近已繳年度,申請人2,申請人3,申請人4,申請人5,FSORT,CNT,cust01,cust02,cust03,cust04,cust05,FCNO From( " & strSql & ") " & _
            " ORDER BY FSort,本所案號"
'end 2023/03/07
CheckOC
adoRecordset.CursorLocation = adUseClient
'Modified by Lydia 2019/11/01 改變型態
'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic

If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
     dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/9/3
     'Added by Lydia 2019/11/01 逐案號判斷
     If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
        intCufaCnt = 0
        adoRecordset.MoveFirst
        Do While adoRecordset.EOF = False
            '利益衝突案件：逐案號判斷
            If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & adoRecordset.Fields("本所案號"), "" & adoRecordset.Fields("cust01") & "," & adoRecordset.Fields("cust02") & "," & adoRecordset.Fields("cust03") & "," & adoRecordset.Fields("cust04") & "," & adoRecordset.Fields("cust05"), "" & adoRecordset.Fields("fcno")) = False Then
                intCufaCnt = intCufaCnt + 1
                adoRecordset.Delete
            End If
            adoRecordset.MoveNext
        Loop
        '利益衝突案件：限閱案件
        If intCufaCnt > 0 Then
            pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
            MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
        End If
        If pub_QL04 <> "" Then InsertQueryLog (dblRow) 'Add By Sindy 2025/9/3
        If adoRecordset.RecordCount = 0 Then
              GoTo JumpToNoData
        End If
     Else
        If pub_QL04 <> "" Then InsertQueryLog (dblRow) 'Add By Sindy 2025/9/3
     End If
    'end 2019/11/01
    Set m_adoRst = adoRecordset.Clone 'Added by Lydia 2018/12/17
Else
   If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/9/3
JumpToNoData:   'Added by Lydia 2019/11/01
   Set m_adoRst = adoRecordset.Clone 'Added by Lydia 2018/02/09
   ShowNoData
   Screen.MousePointer = vbDefault
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If

grdDataList.FixedCols = 0
'Modified by Lydia 2018/12/17 中所反應耗時過久,卡在丟暫存檔(O8的寫法); O12 可以直接排序
'Set GrdDataList.Recordset = adoRecordset
''Added by Lydia 2018/02/09 放到暫存檔,供Grid排序
'Set m_adoRst = PUB_CreateRecordset(adoRecordset, , , 300, Me.Name)
'Modified by Lydia 2018/12/22 拿掉desc
'm_adoRst.Sort = "FSort desc,本所案號 asc" 'Move by Lydia 2018/12/17 先排序,後指定資料集
m_adoRst.Sort = "FSort ,本所案號 asc"
SetRst2Grid
'end 2018/12/17
m_blnColOrderAsc = True
'end 2018/02/09


SetDataListWidth
grdDataList.FixedCols = 4

CheckOC
Me.Enabled = True
End Sub

'add by nickc 2005/09/30 多人申請，沒有只顯示法務的狀況
'Memo frm100102_4用
Sub StrMenu3()
Dim t_pa26 As String
Dim t_pa27 As String
Dim t_pa28 As String
Dim t_pa29 As String
Dim t_pa30 As String
Dim tmpPaCu As Variant
Dim dblRow As Double 'Add By Sindy 2025/9/3

BolFrom100114 = False
Me.Enabled = False
Str01 = ""    '申請人編號
Str02 = ""    '系統類別
Str03 = ""    '收文日期(起)
Str04 = ""    '收文日期(迄)
Str05 = ""    '案件性質(起)
Str06 = ""    '案件性質(迄)
Str07 = ""    '是否含來函資料
Str01 = Me.Tag

SetCustData 'Added by Lydia 2023/08/09 預示顯示客戶編號+名稱
'Add By Sindy 2011/01/03 檢查國內外權限
If CheckSR12(Str01) = False Then
   Screen.MousePointer = vbDefault
   Me.Enabled = True
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If
pub_QL05 = pub_QL05 & ";編號：" & Str01 & "(案件)" 'Add By Sindy 2025/8/13

Str02 = IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys))
If Str02 <> "" Then pub_QL05 = pub_QL05 & ";系統類別：" & Str02 'Add By Sindy 2025/8/13
'收文日起
Str03 = m_Date1
'收文日迄
Str04 = IIf(Len(m_Date1) <= 0, "", IIf(Len(m_Date2) <= 0, (ServerDate - 19110000), m_Date2))
Str05 = m_Pty1
Str06 = m_Pty2
Str07 = m_CKind
'組字串
strSQL1 = ""
If Len(Str03) <> 0 Then
    strSQL1 = strSQL1 + " and cp05>=" & Val(ChangeTStringToWString(Str03))
End If
If Len(Str04) <> 0 Then
    strSQL1 = strSQL1 + " and cp05<=" & Val(ChangeTStringToWString(Str04))
End If
'Add By Sindy 2025/8/13
If Len(Str03) <> 0 Or Len(Str04) <> 0 Then
   pub_QL05 = pub_QL05 & ";收文日期：" & Str03 & "-" & Str04
End If
'2025/8/13 END
If Len(Str05) <> 0 Then
    strSQL1 = strSQL1 + " and cp10>='" & Str05 & "' "
End If
If Len(Str06) <> 0 Then
    strSQL1 = strSQL1 + " and cp10<='" & Str06 & "' "
End If
'Add By Sindy 2025/8/13
If Len(Str05) <> 0 Or Len(Str06) <> 0 Then
   pub_QL05 = pub_QL05 & ";案件性質：" & Str05 & "-" & Str06
End If
'2025/8/13 END
If UCase(Str07) = "N" Then
    strSQL1 = strSQL1 + " and cp09 < 'C' "
    pub_QL05 = pub_QL05 & ";是否含來函資料：不含" 'Add By Sindy 2025/8/13
End If
'顯示表單上面的值
'Modified by Lydya 2023/08/09 模組化
'Label3.Caption = Mid(Me.Tag, 1, InStr(1, Me.Tag, ",") - 1)
'If Len(Trim(Mid(Me.Tag, 1, InStr(1, Me.Tag, ",") - 1))) = 9 Then
'   strSql = "SELECT NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),CU13,ST02,cu111 FROM CUSTOMER,STAFF WHERE CU01='" & Left$(GetNewFagent(Mid(Me.Tag, 1, InStr(1, Me.Tag, ",") - 1)), 8) & "' AND CU02='" & Right$(GetNewFagent(Mid(Me.Tag, 1, InStr(1, Me.Tag, ",") - 1)), 1) & "' AND CU13=ST01(+)"
'Else
'   strSql = "SELECT NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),CU13,ST02,cu111 FROM CUSTOMER,STAFF WHERE CU01='" & Left$(GetNewFagent(Mid(Me.Tag, 1, InStr(1, Me.Tag, ",") - 1)), 8) & "' AND CU02='0' AND CU13=ST01(+) "
'End If
'CheckOC
'adoRecordset.CursorLocation = adUseClient
'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'    If IsNull(adoRecordset.Fields(0)) Then
'        Label4.Caption = ""
'    Else
'        Label4.Caption = adoRecordset.Fields(0)
'    End If
'    If IsNull(adoRecordset.Fields(1)) Then
'        Label6.Caption = ""
'    Else
'        Label6.Caption = adoRecordset.Fields(1)
'    End If
'    If IsNull(adoRecordset.Fields(2)) Then
'        Label7.Caption = ""
'    Else
'        Label7.Caption = adoRecordset.Fields(2)
'    End If
'    If CheckStr(adoRecordset.Fields("cu111")) = "Y" Then
'        Label3.ForeColor = &HFF&
'    Else
'        Label3.ForeColor = &H80000012
'    End If
'End If
'CheckOC
SetCustData
'end 2023/08/09

tmpPaCu = Split(Me.Tag, ",")
t_pa26 = ""
t_pa27 = ""
t_pa28 = ""
t_pa29 = ""
t_pa30 = ""
For i = 0 To UBound(tmpPaCu)
   If tmpPaCu(i) <> "" Then
      Select Case i
      Case 0
               t_pa26 = tmpPaCu(0)
      Case 1
               t_pa27 = tmpPaCu(1)
      Case 2
               t_pa28 = tmpPaCu(2)
      Case 3
               t_pa29 = tmpPaCu(3)
      Case 4
               t_pa30 = tmpPaCu(4)
      Case Else
      End Select
   End If
Next i
         'Added by Lydia 2019/12/26 利益衝突案件：於後面增加欄位
        SeColTM = " ,' ' as cnt,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno "
        SeColPA = " ,' ' as cnt,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno "
        SeColSP = " ,' ' as cnt,sp08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno "
        SeColLC = " ,' ' as cnt,lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno "
        SeColHC = " ,' ' as cnt,hc05 as cust01,hc24 as cust02,hc25 as cust03,hc26 as cust04,hc27 as cust05,'' as fcno "
        'end 2019/12/26
        
        'Add by Amy 2022/12/14 +e符號
        'P、FCP已有專用期間者,或無專用期但進度檔專利之領證601已發文
        'Modify by Amy 2023/03/06 bug-FCP-064401 該案無專用期且無領證發文,不應加e,避免未來資料多跑的慢優化語法
        strSQLE(1) = ",CaseProgress E1"
        strSqlEW(1) = 專利進度註冊費已發文語法
        strEField(1) = "Decode(pa178,'1',Decode(pa24,null,Decode(cp10,'601','e'),'e'))"
        'end 2022/12/14

         'Modify By Sindy 2014/7/7 +||DECODE(PA165,'Y','＃','')
         'Modified by Morgan 2019/1/30 SQLGrpStr(Str02, #)->GetAddStr(str02) 取消不必要系統別的檢查(減少語法執行次數,分所有較大影響)
         'Modified by Lydia 2019/12/26 +增加欄位SeColPA
         'Modify by Amy 2022/12/14 +e符號 Nvl(EState,'')
         'Modify by Amy 2023/03/06 Nvl(EState,'')改為strEField(1),並優化e符號語法
         strSql = " select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','')||" & strEField(1) & " AS 本所案號 ,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
                  "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號, " & cntLstPayYearSQL & "  AS 最近已繳年度,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort" & SeColPA & _
                  " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress " & _
                  " WHERE pa09=na01(+) " & _
                  IIf(Trim(t_pa26) <> "", " and pa26='" & t_pa26 & "' ", " and pa26 is null ") & IIf(Trim(t_pa27) <> "", " and pa27='" & t_pa27 & "' ", " and pa27 is null ") & IIf(Trim(t_pa28) <> "", " and pa28='" & t_pa28 & "' ", " and pa28 is null ") & IIf(Trim(t_pa29) <> "", " and pa29='" & t_pa29 & "' ", " and pa29 is null ") & IIf(Trim(t_pa30) <> "", " and pa30='" & t_pa30 & "' ", " and pa30 is null ") & " and pa01 in (" & GetAddStr(Str02) & ") and pa04='00' " & _
                  " and substr(pa26,1,8)=c1.cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=c1.cu02(+) " & _
                  " and substr(pa27,1,8)=c2.cu01(+) and decode(substr(pa27,9,1),null,'0',substr(pa27,9,1))=c2.cu02(+) " & _
                  " and substr(pa28,1,8)=c3.cu01(+) and decode(substr(pa28,9,1),null,'0',substr(pa28,9,1))=c3.cu02(+) " & _
                  " and substr(pa29,1,8)=c4.cu01(+) and decode(substr(pa29,9,1),null,'0',substr(pa29,9,1))=c4.cu02(+) " & _
                  " and substr(pa30,1,8)=c5.cu01(+) and decode(substr(pa30,9,1),null,'0',substr(pa30,9,1))=c5.cu02(+) " & _
                  " and Pa01=cP01(+) AND Pa02=cP02(+) AND Pa03=cP03(+) AND Pa04=cP04(+) " & strSqlEW(1) & strSQL1
         'Modify By Sindy 2011/1/21 +SP65,SP66
         'Modified by Lydia 2019/12/26 +增加欄位SeColSP
         'Modify by Amy 2020/02/05 +SP73 商品類別
         strSql = strSql + " union select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
                  "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
                  ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort" & SeColSP & _
                  " FROM SERVICEPRACTICE,nation,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5,caseprogress " & _
                  " WHERE sp09=na01(+) " & _
                  IIf(Trim(t_pa26) <> "", " and sp08='" & t_pa26 & "' ", " and sp08 is null ") & IIf(Trim(t_pa27) <> "", " and sp58='" & t_pa27 & "' ", " and sp58 is null ") & IIf(Trim(t_pa28) <> "", " and sp59='" & t_pa28 & "' ", " and sp59 is null ") & IIf(Trim(t_pa29) <> "", " and sp65='" & t_pa29 & "' ", " and sp65 is null ") & IIf(Trim(t_pa30) <> "", " and sp66='" & t_pa30 & "' ", " and sp66 is null ") & " and sp01 in (" & GetAddStr(Str02) & ") AND SP04='00' " & _
                  " AND SUBSTR(SP08,1,8)=C1.CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=C1.CU02(+) " & _
                  " AND SUBSTr(SP58,1,8)=C2.CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=C2.CU02(+) " & _
                  " AND SUBSTR(SP59,1,8)=C3.CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=C3.CU02(+) " & _
                  " AND SUBSTR(SP65,1,8)=C4.CU01(+) AND DECODE(SUBSTR(SP65,9,1),NULL,'0',SUBSTR(SP65,9,1))=C4.CU02(+) " & _
                  " AND SUBSTR(SP66,1,8)=C5.CU01(+) AND DECODE(SUBSTR(SP66,9,1),NULL,'0',SUBSTR(SP66,9,1))=C5.CU02(+) " & _
                  " and sP01=cP01(+) AND sP02=cP02(+) AND sP03=cP03(+) AND sP04=cP04(+) " & strSQL1
         'Modify By Sindy 2014/7/7 +||DECODE(PA165,'Y','＃','')
         '移轉人(讓與人) CP55, CP93~CP96
         'Modified by Lydia 2019/12/26  +增加欄位: CNT, 移轉人(讓與人) 1~5(cust01~cust05), FC代理人(fcno)
         'Modify by Amy 2022/12/14 +e符號 Nvl(EState,'')
         'Modify by Amy 2023/03/06 Nvl(EState,'')改為strEField(1),並優化e符號語法
         strSql = strSql + " union select ' ' AS V,'△'||decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','')||" & Replace(strEField(1), "cp10", "E1.cp10") & " AS 本所案號 , DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , SUBSTRB(NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)),1,10) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
                  "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度" & _
                  ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort,' ' as cnt " & _
                  ",CP.cp55 as cust01,CP.cp93 as cust02,CP.cp94 as cust03,CP.cp95 as cust04,CP.cp96 as cust05,pa75 as fcno" & _
                  " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress CP" & strSQLE(1) & " " & _
                  " WHERE pa09=na01(+) " & _
                  IIf(Trim(t_pa26) <> "", " and CP.cp55='" & t_pa26 & "' ", " and CP.cp55 is null ") & IIf(Trim(t_pa27) <> "", " and CP.cp93='" & t_pa27 & "' ", " and CP.cp93 is null ") & IIf(Trim(t_pa28) <> "", " and CP.cp94='" & t_pa28 & "' ", " and CP.cp94 is null ") & IIf(Trim(t_pa29) <> "", " and CP.cp95='" & t_pa29 & "' ", " and CP.cp95 is null ") & IIf(Trim(t_pa30) <> "", " and CP.cp96='" & t_pa30 & "' ", " and CP.cp96 is null ") & " and pa01 in (" & GetAddStr(Str02) & ") and pa04='00' " & _
                  " and substr(CP.CP55,1,8)=c1.cu01(+) and decode(substr(CP.CP55,9,1),null,'0',substr(CP.CP55,9,1))=c1.cu02(+) " & _
                  " and substr(CP.cp93,1,8)=c2.cu01(+) and decode(substr(CP.cp93,9,1),null,'0',substr(CP.cp93,9,1))=c2.cu02(+) " & _
                  " and substr(CP.cp94,1,8)=c3.cu01(+) and decode(substr(CP.cp94,9,1),null,'0',substr(CP.cp94,9,1))=c3.cu02(+) " & _
                  " and substr(CP.cp95,1,8)=c4.cu01(+) and decode(substr(CP.cp95,9,1),null,'0',substr(CP.cp95,9,1))=c4.cu02(+) " & _
                  " and substr(CP.cp96,1,8)=c5.cu01(+) and decode(substr(CP.cp96,9,1),null,'0',substr(CP.cp96,9,1))=c5.cu02(+) " & _
                  " and CP.CP01=PA01(+) AND CP.CP02=PA02(+) AND CP.CP03=PA03(+) AND CP.CP04=PA04(+) " & Replace(strSqlEW(1), "cp", "E1.cp") & strSQL1
                  
         '移轉申請人(讓與申請人) CP56,CP89~CP92
         'Modified by Lydia 2019/12/26  +增加欄位:CNT, 移轉申請人(讓與申請人) 1~5(cust01~cust05), FC代理人(fcno)
         'Modify by Amy 2022/12/14 +e符號 Nvl(EState,'')
         'Modify by Amy 2023/03/06 Nvl(EState,'')改為strEField(1),並優化e符號語法
         strSql = strSql + " union select ' ' AS V,'△'||decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||DECODE(PA165,'Y','＃','')||" & Replace(strEField(1), "cp10", "E1.cp10") & " AS 本所案號 , DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , SUBSTRB(NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)),1,10) AS 申請人1 ,' ' AS 商品類別,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||" & _
                  "DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間, NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度" & _
                  ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort,' ' as cnt " & _
                  ",CP.cp56 as cust01,CP.cp89 as cust02,CP.cp90 as cust03,CP.cp91 as cust04,CP.cp92 as cust05,pa75 as fcno" & _
                  " FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5, caseprogress CP" & strSQLE(1) & " " & _
                  " WHERE pa09=na01(+) " & _
                  IIf(Trim(t_pa26) <> "", " and CP.cp55='" & t_pa26 & "' ", " and CP.cp55 is null ") & IIf(Trim(t_pa27) <> "", " and CP.cp93='" & t_pa27 & "' ", " and CP.cp93 is null ") & IIf(Trim(t_pa28) <> "", " and CP.cp94='" & t_pa28 & "' ", " and CP.cp94 is null ") & IIf(Trim(t_pa29) <> "", " and CP.cp95='" & t_pa29 & "' ", " and CP.cp95 is null ") & IIf(Trim(t_pa30) <> "", " and CP.cp96='" & t_pa30 & "' ", " and CP.cp96 is null ") & " and pa01 in (" & GetAddStr(Str02) & ") and pa04='00' " & _
                  " and substr(CP.CP56,1,8)=c1.cu01(+) and decode(substr(CP.CP56,9,1),null,'0',substr(CP.CP56,9,1))=c1.cu02(+) " & _
                  " and substr(CP.cp89,1,8)=c2.cu01(+) and decode(substr(CP.cp89,9,1),null,'0',substr(CP.cp89,9,1))=c2.cu02(+) " & _
                  " and substr(CP.cp90,1,8)=c3.cu01(+) and decode(substr(CP.cp90,9,1),null,'0',substr(CP.cp90,9,1))=c3.cu02(+) " & _
                  " and substr(CP.cp91,1,8)=c4.cu01(+) and decode(substr(CP.cp91,9,1),null,'0',substr(CP.cp91,9,1))=c4.cu02(+) " & _
                  " and substr(CP.cp92,1,8)=c5.cu01(+) and decode(substr(CP.cp92,9,1),null,'0',substr(CP.cp92,9,1))=c5.cu02(+) " & _
                  " and CP.CP01=PA01(+) AND CP.CP02=PA02(+) AND CP.CP03=PA03(+) AND CP.CP04=PA04(+) " & Replace(strSqlEW(1), "cp", "E1.cp") & strSQL1
                  
         '移轉人(讓與人) CP55, CP93~CP96
         'Modified by Lydia 2019/11/01  +增加欄位:CNT, 移轉人(讓與人) 1~5(cust01~cust05), FC代理人(fcno)
         'Modify by Amy 2020/02/05 +SP73 商品類別
         strSql = strSql + " union select ' ' AS V,'△'||SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , SUBSTRB(NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)),1,10) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
                  "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
                  ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort,' ' as cnt " & _
                  ",cp55 as cust01,cp93 as cust02,cp94 as cust03,cp95 as cust04,cp96 as cust05,sp26 as fcno" & _
                  " FROM SERVICEPRACTICE,nation,customer c1,customer c2,customer c3,customer c4,customer c5,caseprogress " & _
                  " WHERE sp09=na01(+) " & _
                  IIf(Trim(t_pa26) <> "", " and cp55='" & t_pa26 & "' ", " and cp55 is null ") & IIf(Trim(t_pa27) <> "", " and cp93='" & t_pa27 & "' ", " and cp93 is null ") & IIf(Trim(t_pa28) <> "", " and cp94='" & t_pa28 & "' ", " and cp94 is null ") & IIf(Trim(t_pa29) <> "", " and cp95='" & t_pa29 & "' ", " and cp95 is null ") & IIf(Trim(t_pa30) <> "", " and cp96='" & t_pa30 & "' ", " and cp96 is null ") & " and sp01 in (" & GetAddStr(Str02) & ") and sp04='00' " & _
                  " and substr(CP55,1,8)=c1.cu01(+) and decode(substr(CP55,9,1),null,'0',substr(CP55,9,1))=c1.cu02(+) " & _
                  " and substr(cp93,1,8)=c2.cu01(+) and decode(substr(cp93,9,1),null,'0',substr(cp93,9,1))=c2.cu02(+) " & _
                  " and substr(cp94,1,8)=c3.cu01(+) and decode(substr(cp94,9,1),null,'0',substr(cp94,9,1))=c3.cu02(+) " & _
                  " and substr(cp95,1,8)=c4.cu01(+) and decode(substr(cp95,9,1),null,'0',substr(cp95,9,1))=c4.cu02(+) " & _
                  " and substr(cp96,1,8)=c5.cu01(+) and decode(substr(cp96,9,1),null,'0',substr(cp96,9,1))=c5.cu02(+) " & _
                  " and CP01=sP01(+) AND CP02=sP02(+) AND CP03=sP03(+) AND CP04=sP04(+) " & strSQL1
         
         '移轉申請人(讓與申請人) CP56,CP89~CP92
         'Modified by Lydia 2019/11/01  +增加欄位:CNT, 移轉申請人(讓與申請人) 1~5(cust01~cust05), FC代理人(fcno)
         'Modify by Amy 2020/02/05 +SP73 商品類別
         strSql = strSql + " union select ' ' AS V,'△'||SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號 , DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號 ,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,NVL(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , SUBSTRB(NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)),1,10) AS 申請人1 ,NVL(SP73,'') AS 商品類別,DECODE(SP20,NULL,'','','',(SUBSTR(SP20,1,4)||'/'||SUBSTR(SP20,5,2)||'/'||SUBSTR(SP20,7,2)))||'-'||" & _
                  "DECODE(SP21,NULL,'','','',(SUBSTR(SP21,1,4)||'/'||SUBSTR(SP21,5,2)||'/'||SUBSTR(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度" & _
                  ",NVL(C2.CU04,nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,nvl(C3.CU04,nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort,' ' as cnt " & _
                  ",cp56 as cust01,cp89 as cust02,cp90 as cust03,cp91 as cust04,cp92 as cust05,sp26 as fcno" & _
                  " FROM SERVICEPRACTICE,nation,customer c1,customer c2,customer c3,customer c4,customer c5, caseprogress " & _
                  " WHERE sp09=na01(+) " & _
                  IIf(Trim(t_pa26) <> "", " and cp55='" & t_pa26 & "' ", " and cp55 is null ") & IIf(Trim(t_pa27) <> "", " and cp93='" & t_pa27 & "' ", " and cp93 is null ") & IIf(Trim(t_pa28) <> "", " and cp94='" & t_pa28 & "' ", " and cp94 is null ") & IIf(Trim(t_pa29) <> "", " and cp95='" & t_pa29 & "' ", " and cp95 is null ") & IIf(Trim(t_pa30) <> "", " and cp96='" & t_pa30 & "' ", " and cp96 is null ") & " and sp01 in (" & GetAddStr(Str02) & ") and sp04='00' " & _
                  " and substr(CP56,1,8)=c1.cu01(+) and decode(substr(CP56,9,1),null,'0',substr(CP56,9,1))=c1.cu02(+) " & _
                  " and substr(cp89,1,8)=c2.cu01(+) and decode(substr(cp89,9,1),null,'0',substr(cp89,9,1))=c2.cu02(+) " & _
                  " and substr(cp90,1,8)=c3.cu01(+) and decode(substr(cp90,9,1),null,'0',substr(cp90,9,1))=c3.cu02(+) " & _
                  " and substr(cp91,1,8)=c4.cu01(+) and decode(substr(cp91,9,1),null,'0',substr(cp91,9,1))=c4.cu02(+) " & _
                  " and substr(cp92,1,8)=c5.cu01(+) and decode(substr(cp92,9,1),null,'0',substr(cp92,9,1))=c5.cu02(+) " & _
                  " and CP01=sP01(+) AND CP02=sP02(+) AND CP03=sP03(+) AND CP04=sP04(+) " & strSQL1

If m_Cty1 <> "" Or m_Cty2 <> "" Then
   strSql = "Select X.* From (" & strSql & ") X,Nation Y Where na03(+)=申請國家"
   If m_Cty1 <> "" Then
      strSql = strSql & " and na01>='" & m_Cty1 & "'"
   End If
   If m_Cty2 <> "" Then
      strSql = strSql & " and na01<='" & m_Cty2 & "'"
   End If
End If

strSql = strSql & " ORDER BY FSort,本所案號"

CheckOC
adoRecordset.CursorLocation = adUseClient
'Modified by Lydia 2019/11/01 改變型態
'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic

If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/9/3
    'Remove by Lydia 2019/11/01
    'If Len(Trim(Str02)) <> 0 Then
    '    strTemp = Split(Str02, ",")
    'End If
    'adoRecordset.MoveFirst
    'Dim StrTest2 As String, StrTest4 As String, s As Integer
    
     'Added by Lydia 2019/11/01 逐案號判斷
     If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
        intCufaCnt = 0
        adoRecordset.MoveFirst
        Do While adoRecordset.EOF = False
            '利益衝突案件：逐案號判斷
            If PUB_ChkCufaByCase(Me.Name, Str02, "" & adoRecordset.Fields("本所案號"), "" & adoRecordset.Fields("cust01") & "," & adoRecordset.Fields("cust02") & "," & adoRecordset.Fields("cust03") & "," & adoRecordset.Fields("cust04") & "," & adoRecordset.Fields("cust05"), "" & adoRecordset.Fields("fcno")) = False Then
                intCufaCnt = intCufaCnt + 1
                adoRecordset.Delete
            End If
            adoRecordset.MoveNext
        Loop
        '利益衝突案件：限閱案件
        If intCufaCnt > 0 Then
            pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
            MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
        End If
        If pub_QL04 <> "" Then InsertQueryLog (dblRow)  'Add By Sindy 2025/9/3
        If adoRecordset.RecordCount = 0 Then
              GoTo JumpToNoData
        End If
     Else
        If pub_QL04 <> "" Then InsertQueryLog (dblRow)  'Add By Sindy 2025/9/3
     End If
    'end 2019/11/01
    
    Set m_adoRst = adoRecordset.Clone 'Added by Lydia 2018/02/09 'move by Lydia 2018/12/17 從下面移上來
    If adoRecordset.RecordCount = 0 Then
        Me.Enabled = True
        cmdok(0).Enabled = False
        cmdok(1).Enabled = False
        cmdok(2).Enabled = False
        cmdok(3).Enabled = False
        ShowNoData
        Screen.MousePointer = vbDefault
        Me.Enabled = True
       tmpBol = fnCancelNowFormAndShowParentForm(Me)

        Exit Sub
    End If
Else
   If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/9/3
JumpToNoData:   'Added by Lydia 2019/11/01
   Set m_adoRst = adoRecordset.Clone 'Added by Lydia 2018/02/09
   ShowNoData
   Screen.MousePointer = vbDefault
   Me.Enabled = True
     tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If

grdDataList.FixedCols = 0
'Modified by Lydia 2018/12/17 中所反應耗時過久,卡在丟暫存檔(O8的寫法); O12 可以直接排序
'Set GrdDataList.Recordset = adoRecordset
''Added by Lydia 2018/02/09 放到暫存檔,供Grid排序
'Set m_adoRst = PUB_CreateRecordset(adoRecordset, , , 300, Me.Name)
'Modified by Lydia 2018/12/22 拿掉desc
'm_adoRst.Sort = "FSort desc,本所案號 asc" 'Move by Lydia 2018/12/17 先排序,後指定資料集
m_adoRst.Sort = "FSort ,本所案號 asc"
SetRst2Grid
'end 2018/12/17
m_blnColOrderAsc = True
'end 2018/02/09

SetDataListWidth
grdDataList.FixedCols = 4

Me.Enabled = True
End Sub

'Add by Amy 2014/05/07 由對造資料過來的
Sub StrMenu4()
Dim dblRow As Double 'Add By Sindy 2025/9/3

BolFrom100114 = False
Me.Enabled = False
Str01 = ""    '申請人編號
Str02 = ""    '系統別
Str03 = ""    '本所案號
Str01 = Me.Tag
Str03 = m_CaseNo
SetCustData 'Added by Lydia 2023/08/09 預示顯示客戶編號+名稱

'檢查國內外權限
If CheckSR12(Str01) = False Then
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    tmpBol = fnCancelNowFormAndShowParentForm(Me)
    Exit Sub
End If
pub_QL05 = pub_QL05 & ";編號：" & Str01 & "(案件)" 'Add By Sindy 2025/8/13

'Modifed by Lydia 2023/08/09 模組化
'Label3.Caption = Me.Tag
'If Len(Trim(Me.Tag)) = 9 Then
'   strSql = "SELECT NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),CU13,ST02,cu111 FROM CUSTOMER,STAFF WHERE CU01='" & Left$(GetNewFagent(Me.Tag), 8) & "' AND CU02='" & Right$(GetNewFagent(Me.Tag), 1) & "' AND CU13=ST01(+)"
'Else
'   strSql = "SELECT NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),CU13,ST02,cu111 FROM CUSTOMER,STAFF WHERE CU01='" & Left$(GetNewFagent(Me.Tag), 8) & "' AND CU02='0' AND CU13=ST01(+) "
'End If
'CheckOC
'adoRecordset.CursorLocation = adUseClient
'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'    If IsNull(adoRecordset.Fields(0)) Then
'        Label4.Caption = ""
'    Else
'        Label4.Caption = adoRecordset.Fields(0)
'    End If
'    If IsNull(adoRecordset.Fields(1)) Then
'        Label6.Caption = ""
'    Else
'        Label6.Caption = adoRecordset.Fields(1)
'    End If
'    If IsNull(adoRecordset.Fields(2)) Then
'        Label7.Caption = ""
'    Else
'        Label7.Caption = adoRecordset.Fields(2)
'    End If
'    If CheckStr(adoRecordset.Fields("cu111")) = "Y" Then
'        Label3.ForeColor = &HFF&
'    Else
'        Label3.ForeColor = &H80000012
'    End If
'End If
'CheckOC
SetCustData
'end 2023/08/09

 'Added by Lydia 2019/12/26 利益衝突案件：於後面增加欄位
 SeColTM = " ,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno "
 SeColPA = " ,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno "
 SeColSP = " ,sp08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno "
 SeColLC = " ,lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno "
 SeColHC = " ,hc05 as cust01,hc24 as cust02,hc25 as cust03,hc26 as cust04,hc27 as cust05,'' as fcno "
 'end 2019/12/26
 
'Add by Amy 2022/12/14 +e符號
'Modify by Amy 2023/03/06 bug-FCP-064401 該案無專用期且無領證發文,不應加e,避免未來資料多跑的慢優化語法
'T、FCT已有專用期間者,或無專用期但進度檔商標之註冊費717已發文
strSQLE(0) = ",CaseProgress "
strSqlEW(0) = Replace(Replace(專利進度註冊費已發文語法, "601", "717"), "pa", "tm")
strEField(0) = "Decode(tm136,'1',Decode(tm21,null,Decode(cp10,'717','e'),'e'))"
'P、FCP已有專用期間者,或無專用期但進度檔專利之領證601已發文
strSQLE(1) = ",CaseProgress "
strSqlEW(1) = 專利進度註冊費已發文語法
strEField(1) = "Decode(pa178,'1',Decode(pa24,null,Decode(cp10,'601','e'),'e'))"
'end 2023/03/06
'end 2022/12/14
 
Str02 = SystemNumber(Str03, 1)
pub_QL05 = pub_QL05 & ";本所案號：" & Str03 'Add By Sindy 2025/8/13
Str03 = Replace(m_CaseNo, "-", "")
Select Case Str02
    ' 讀取商標基本檔
    Case "T", "TF", "CFT", "FCT":
        'Modified by Lydia 2019/12/26 +增加欄位SeColTM
        'Modify by Amy 2022/12/14 +e符號 Nvl(EState,'')
        'Modify by Amy 2023/03/06 Nvl(EState,'')改為strEField(1),並優化e符號語法
        strExc(0) = "Select ' ' AS V,Decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||Decode(TM29,'Y','＊','')||Decode(Length(Nvl(tm57,'')),Null,'','●')||Decode(TM128,'Y','◎','A','□','')||" & strEField(0) & " AS 本所案號 ,Decode(Length(Nvl(tm73,'')),Null,'','●')||tm34 AS 分所號 ,Nvl(TM05,Nvl(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定專利號數,Decode(TM16,'1','准','2','駁',' ') AS 准駁 , Nvl(C1.CU04,Nvl(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,Nvl(TM09,' ') AS 商品類別," & _
                            "Decode(TM21,Null,'','','',(SubStr(TM21,1,4)||'/'||SubStr(TM21,5,2)||'/'||SubStr(TM21,7,2)))||'-'||Decode(TM22,Null,'','','',(SubStr(TM22,1,4)||'/'||SubStr(TM22,5,2)||'/'||SubStr(TM22,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,Nvl(C2.CU04,Nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,Nvl(C3.CU04,Nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,Nvl(C4.CU04,Nvl(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,Nvl(C5.CU04,Nvl(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5," & _
                            "TM01||'-'||TM02||'-'||TM03||'-'||TM04||Decode(TM29,'Y','＊','')||Decode(Length(Nvl(tm57,'')),Null,'','●')||Decode(TM128,'Y','◎','A','□','') as FSort,Decode(TM123,Null,C1.CU01||C1.CU127,C1.CU01||TM123) CNT " & SeColTM & _
                            "From TRADEMARK,nation,Customer c1,Customer c2,Customer c3,Customer c4,Customer c5" & strSQLE(0) & " " & _
                            "Where " & ChgTradeMark(Str03) & " And tm10=na01(+) And SubStr(TM23,1,8) = c1.CU01(+) And Decode(SubStr(TM23,9,1),Null,'0',SubStr(TM23,9,1)) = c1.CU02(+) " & _
                            "And SubStr(tm78,1,8)=c2.cu01(+) And Decode(SubStr(tm78,9,1),Null,'0',SubStr(tm78,9,1))=c2.cu02(+) " & _
                            "And SubStr(tm79,1,8)=c3.cu01(+) And Decode(SubStr(tm79,9,1),Null,'0',SubStr(tm79,9,1))=c3.cu02(+) " & _
                            "And SubStr(tm80,1,8)=c4.cu01(+) And Decode(SubStr(tm80,9,1),Null,'0',SubStr(tm80,9,1))=c4.cu02(+) " & _
                            "And SubStr(tm81,1,8)=c5.cu01(+) And Decode(SubStr(tm81,9,1),Null,'0',SubStr(tm81,9,1))=c5.cu02(+) " & strSqlEW(0)
        
    ' 讀取專利基本檔
    Case "P", "CFP", "FCP":
        'Modify By Sindy 2014/7/7 +||DECODE(PA165,'Y','＃','')
        'Modified by Lydia 2019/12/26 +增加欄位SeColPA
        'Modify by Amy 2022/12/14 +e符號 Nvl(EState,'')
        'Modify by Amy 2023/03/06 Nvl(EState,'')改為strEField(1),並優化e符號語法
        strExc(0) = "Select ' ' AS V,Decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||Decode(PA57,'Y','＊','')||Decode(Length(Nvl(pa108,'')),Null,'','●')||DECODE(PA165,'Y','＃','')||" & strEField(1) & " AS 本所案號 ,Decode(Length(Nvl(pa136,'')),Null,'','●')||pa47 AS 分所號 ,Nvl(PA05,Nvl(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 審定專利號數,Decode(PA16,'1','准','2','駁',' ') AS 准駁 , Nvl(C1.CU04,Nvl(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,' ' AS 商品類別," & _
                            "Decode(PA24,Null,'','','',(SubStr(PA24,1,4)||'/'||SubStr(PA24,5,2)||'/'||SubStr(PA24,7,2)))||'-'||Decode(PA25,Null,'','','',(SubStr(PA25,1,4)||'/'||SubStr(PA25,5,2)||'/'||SubStr(PA25,7,2))) AS 專用期間, Nvl(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度,Nvl(C2.CU04,Nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,Nvl(C3.CU04,Nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,Nvl(C4.CU04,Nvl(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,Nvl(C5.CU04,Nvl(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5," & _
                            "PA01||'-'||PA02||'-'||PA03||'-'||PA04||Decode(PA57,'Y','＊','')||Decode(Length(Nvl(pa108,'')),Null,'','●') as FSort,Decode(PA149,Null,C1.CU01||C1.CU127,C1.CU01||PA149) CNT " & SeColPA & _
                            "From PATENT,nation,Customer c1,Customer c2,Customer c3,Customer c4,Customer c5" & strSQLE(1) & " " & _
                            "Where " & ChgPatent(Str03) & " And pa09=na01(+) And SubStr(pa26,1,8)=c1.cu01(+) And Decode(SubStr(pa26,9,1),Null,'0',SubStr(pa26,9,1))=c1.cu02(+) " & _
                            "And SubStr(pa27,1,8)=c2.cu01(+) And Decode(SubStr(pa27,9,1),Null,'0',SubStr(pa27,9,1))=c2.cu02(+) " & _
                            "And SubStr(pa28,1,8)=c3.cu01(+) And Decode(SubStr(pa28,9,1),Null,'0',SubStr(pa28,9,1))=c3.cu02(+) " & _
                            "And SubStr(pa29,1,8)=c4.cu01(+) And Decode(SubStr(pa29,9,1),Null,'0',SubStr(pa29,9,1))=c4.cu02(+) " & _
                            "And SubStr(pa30,1,8)=c5.cu01(+) And Decode(SubStr(pa30,9,1),Null,'0',SubStr(pa30,9,1))=c5.cu02(+) " & strSqlEW(1)
                            
    ' 讀取法務基本檔
    'modify by sonia 2019/7/29 +ACS系統類別
    Case "L", "CFL", "FCL", "LIN", "ACS":
        'Modified by Lydia 2019/12/26 +增加欄位SeColLC
        strExc(0) = "Select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||Decode(LC08,'Y','＊','')||Decode(Length(Nvl(lc34,'')),Null,'','●') AS 本所案號 , Decode(Length(Nvl(lc36,'')),Null,'','●')||lc16 AS 分所號 ,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,na03 AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , Nvl(c1.CU04,Nvl(c1.CU05||c1.CU88||c1.CU89||c1.CU90,c1.CU06)) AS 申請人1 ,' ' AS 商品類別," & _
                            "'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,Nvl(c2.CU04,Nvl(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2, Nvl(c3.CU04,Nvl(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3, Nvl(c4.CU04,Nvl(c4.CU05||c4.CU88||c4.CU89||c4.CU90,c4.CU06)) AS 申請人4, Nvl(c5.CU04,Nvl(c5.CU05||c5.CU88||c5.CU89||c5.CU90,c5.CU06)) AS 申請人5," & _
                            "LC01||'-'||LC02||'-'||LC03||'-'||LC04||Decode(LC08,'Y','＊','')||Decode(Length(Nvl(lc34,'')),Null,'','●') as FSort,Decode(LC42,Null,c1.CU01||c1.CU127,c1.CU01||LC42) CNT " & SeColLC & _
                            "From LAWCASE,nation,Customer c1,Customer c2,Customer c3,Customer c4,Customer c5 " & _
                            "Where " & ChgLawcase(Str03) & "And lc15=na01(+) And SubStr(LC11,1,8)=c1.CU01(+) And Decode(SubStr(LC11,9,1),Null,'0',SubStr(LC11,9,1)) = c1.CU02(+) " & _
                            "And SubStr(lc43,1,8)=C2.CU01(+) And Decode(SubStr(lc43,9,1),Null,'0',SubStr(lc43,9,1))=C2.CU02(+) " & _
                            "And SubStr(lc44,1,8)=C3.CU01(+) And Decode(SubStr(lc44,9,1),Null,'0',SubStr(lc44,9,1))=C3.CU02(+) " & _
                            "And SubStr(lc45,1,8)=C4.CU01(+) And Decode(SubStr(lc45,9,1),Null,'0',SubStr(lc45,9,1))=C4.CU02(+) " & _
                            "And SubStr(lc46,1,8)=C5.CU01(+) And Decode(SubStr(lc46,9,1),Null,'0',SubStr(lc46,9,1))=C5.CU02(+) "
            
    ' 讀取顧問案件基本檔
    Case "LA":
        'Modified by Lydia 2019/12/26 +增加欄位SeColHC
        strExc(0) = "Select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||Decode(HC09,'Y','＊','')||Decode(Length(Nvl(hc19,'')),Null,'','●') AS 本所案號 , Decode(Length(Nvl(hc20,'')),Null,'','●')||hc07 AS 分所號 ,HC06 AS 案件名稱,'台灣' AS 申請國家,' ' AS 申請案號 , ' ' AS 申請日 ,' ' AS 審定專利號數,' ' AS 准駁 , Nvl(c1.CU04,Nvl(c1.CU05||c1.CU88||c1.CU89||c1.CU90,c1.CU06)) AS 申請人1 ,' ' AS 商品類別," & _
                            "'-' AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,Nvl(c2.CU04,Nvl(c2.CU05||c2.CU88||c2.CU89||c2.CU90,c2.CU06)) AS 申請人2, Nvl(c3.CU04,Nvl(c3.CU05||c3.CU88||c3.CU89||c3.CU90,c3.CU06)) AS 申請人3, Nvl(c4.CU04,Nvl(c4.CU05||c4.CU88||c4.CU89||c4.CU90,c4.CU06)) AS 申請人4, Nvl(c5.CU04,Nvl(c5.CU05||c5.CU88||c5.CU89||c5.CU90,c5.CU06)) AS 申請人5," & _
                            "HC01||'-'||HC02||'-'||HC03||'-'||HC04||Decode(HC09,'Y','＊','')||Decode(Length(Nvl(hc19,'')),Null,'','●')  as FSort,Decode(HC23,Null,c1.CU01||c1.CU127,c1.CU01||HC23) CNT " & SeColHC & _
                            "From HIRECASE,Customer c1,Customer c2,Customer c3,Customer c4,Customer c5 " & _
                            "Where " & ChgHirecase(Str03) & " And SubStr(HC05,1,8)=c1.CU01(+) And Decode(SubStr(HC05,9,1),Null,'0',SubStr(HC05,9,1))=c1.CU02(+) " & _
                            "And SubStr(hc24,1,8)=C2.CU01(+) And Decode(SubStr(hc24,9,1),Null,'0',SubStr(hc24,9,1))=C2.CU02(+) " & _
                            "And SubStr(hc25,1,8)=C3.CU01(+) And Decode(SubStr(hc25,9,1),Null,'0',SubStr(hc25,9,1))=C3.CU02(+) " & _
                            "And SubStr(hc26,1,8)=C4.CU01(+) And Decode(SubStr(hc26,9,1),Null,'0',SubStr(hc26,9,1))=C4.CU02(+) " & _
                            "And SubStr(hc27,1,8)=C5.CU01(+) And Decode(SubStr(hc27,9,1),Null,'0',SubStr(hc27,9,1))=C5.CU02(+) "
            
    ' 讀取服務業務基本檔
    Case Else:
        'Modified by Lydia 2019/12/26 +增加欄位SeColSP
        'Moidfy by Amy 2020/02/05 +SP73 商品類別
        strExc(0) = "Select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||Decode(sp15,'Y','＊','')||Decode(Length(Nvl(sp61,'')),Null,'','●') AS 本所案號 , Decode(Length(Nvl(sp68,'')),Null,'','●')||sp28 AS 分所號 ,Nvl(SP05,Nvl(SP06,SP07)) AS 案件名稱,na03 AS 申請國家,SP11 AS 申請案號 ," & SQLDate("SP10") & " AS 申請日 ,Nvl(SP14,SP13) AS 審定專利號數,' ' AS 准駁 , Nvl(C1.CU04,Nvl(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(SP73,'') AS 商品類別," & _
                            "Decode(SP20,Null,'','','',(SubStr(SP20,1,4)||'/'||SubStr(SP20,5,2)||'/'||SubStr(SP20,7,2)))||'-'||Decode(SP21,Null,'','','',(SubStr(SP21,1,4)||'/'||SubStr(SP21,5,2)||'/'||SubStr(SP21,7,2))) AS 專用期間, ' '  AS 專利公告號, ' ' AS 最近已繳年度,Nvl(C2.CU04,Nvl(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,Nvl(C3.CU04,Nvl(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,Nvl(C4.CU04,Nvl(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,Nvl(C5.CU04,Nvl(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5," & _
                            "SP01||'-'||SP02||'-'||SP03||'-'||SP04||Decode(sp15,'Y','＊','')||Decode(Length(Nvl(sp61,'')),Null,'','●')  as FSort,Decode(SP78,Null,C1.CU01||C1.CU127,C1.CU01||SP78) CNT " & SeColSP & _
                            "From SERVICEPRACTICE,nation,Customer C1,Customer C2,Customer C3,Customer c4,Customer c5 " & _
                            "Where " & ChgService(Str03) & " And sp09=na01(+) And SubStr(SP08,1,8)=C1.CU01(+) And Decode(SubStr(SP08,9,1),Null,'0',SubStr(SP08,9,1))=C1.CU02(+) " & _
                            "And SubStr(SP58,1,8)=C2.CU01(+) And Decode(SubStr(SP58,9,1),Null,'0',SubStr(SP58,9,1))=C2.CU02(+) " & _
                            "And SubStr(SP59,1,8)=C3.CU01(+) And Decode(SubStr(SP59,9,1),Null,'0',SubStr(SP59,9,1))=C3.CU02(+) " & _
                            "And SubStr(sp65,1,8)=c4.cu01(+) And Decode(SubStr(sp65,9,1),Null,'0',SubStr(sp65,9,1))=c4.cu02(+) " & _
                            "And SubStr(sp66,1,8)=c5.cu01(+) And Decode(SubStr(sp66,9,1),Null,'0',SubStr(sp66,9,1))=c5.cu02(+) "
                             
End Select
    
If ChkPCT.Value = vbChecked Then
    strSql = Replace(Replace(Replace(Replace(UCase(strSql), "TM15 AS 審定專利號數", "'' as PCT"), "NVL(SP14,SP13) AS 審定專利號數", "'' as PCT"), "' ' AS 審定專利號數", "'' as PCT"), "PA22 AS 審定專利號數", "pa46 as PCT")
End If
strExc(0) = strExc(0) & " Order by FSort,本所案號"
    
CheckOC
adoRecordset.CursorLocation = adUseClient
'Modified by Lydia 2019/11/01 改變型態
'adoRecordset.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
adoRecordset.Open strExc(0), cnnConnection, adOpenDynamic, adLockBatchOptimistic
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/9/3
   'Remove by Lydia 2019/11/01
   'If Len(Trim(Str02)) <> 0 Then
   '    strTemp = Split(Str02, ",")
   'End If
   'adoRecordset.MoveFirst
   
   'Added by Lydia 2019/11/01 逐案號判斷
   If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
      intCufaCnt = 0
      adoRecordset.MoveFirst
      Do While adoRecordset.EOF = False
          '利益衝突案件：逐案號判斷
          If PUB_ChkCufaByCase(Me.Name, Str02, "" & adoRecordset.Fields("本所案號"), "" & adoRecordset.Fields("cust01") & "," & adoRecordset.Fields("cust02") & "," & adoRecordset.Fields("cust03") & "," & adoRecordset.Fields("cust04") & "," & adoRecordset.Fields("cust05"), "" & adoRecordset.Fields("fcno")) = False Then
              intCufaCnt = intCufaCnt + 1
              adoRecordset.Delete
          End If
          adoRecordset.MoveNext
      Loop
      '利益衝突案件：限閱案件
      If intCufaCnt > 0 Then
         pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
         MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
      End If
      If pub_QL04 <> "" Then InsertQueryLog (dblRow) 'Add By Sindy 2025/8/13
      If adoRecordset.RecordCount = 0 Then
         GoTo JumpToNoData
      End If
   Else
      If pub_QL04 <> "" Then InsertQueryLog (dblRow) 'Add By Sindy 2025/8/13
   End If
   'end 2019/11/01
   
   Set m_adoRst = adoRecordset.Clone 'Added by Lydia 2018/02/09 'move by Lydia 2018/12/17 從下面移上來
   If adoRecordset.RecordCount = 0 Then
       Me.Enabled = True
       cmdok(0).Enabled = False
       cmdok(1).Enabled = False
       cmdok(2).Enabled = False
       cmdok(3).Enabled = False
       ShowNoData
       Screen.MousePointer = vbDefault
       Me.Enabled = True
       tmpBol = fnCancelNowFormAndShowParentForm(Me)
       Exit Sub
   End If
    
Else
   If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/13
JumpToNoData:   'Added by Lydia 2019/11/01
   Set m_adoRst = adoRecordset.Clone 'Added by Lydia 2018/02/09
   ShowNoData
   Screen.MousePointer = vbDefault
   Me.Enabled = True
   tmpBol = fnCancelNowFormAndShowParentForm(Me)

   Exit Sub
End If

grdDataList.FixedCols = 0
'Modified by Lydia 2018/12/17 中所反應耗時過久,卡在丟暫存檔(O8的寫法); O12 可以直接排序
'Set GrdDataList.Recordset = adoRecordset
''Added by Lydia 2018/02/09 放到暫存檔,供Grid排序
'Set m_adoRst = PUB_CreateRecordset(adoRecordset, , , 300, Me.Name)
'Modified by Lydia 2018/12/22 拿掉desc
'm_adoRst.Sort = "FSort desc,本所案號 asc" 'Move by Lydia 2018/12/17 先排序,後指定資料集
m_adoRst.Sort = "FSort ,本所案號 asc"
SetRst2Grid
'end 2018/12/17
m_blnColOrderAsc = True
'end 2018/02/09

SetDataListWidth
grdDataList.FixedCols = 4
Me.Enabled = True

cmdok(8).Enabled = False
cmdok(8).Visible = False
End Sub

'Added by Lydia 2018/02/09 設暫存檔為Grid來源
Private Sub SetRst2Grid()
   grdDataList.FixedCols = 0
   Set grdDataList.Recordset = m_adoRst
    If bolIsL = False Then
       grdDataList.FixedCols = 4
    Else
       grdDataList.FixedCols = 6
    End If
End Sub

'Added by Lydia 2018/02/09
Private Sub grdDataList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim iCol As Integer
    iCol = grdDataList.MouseCol
    If grdDataList.MouseRow < 1 Then
      grdDataList.Visible = False
      Set grdDataList.Recordset = Nothing
      If m_blnColOrderAsc = True Then
         m_adoRst.Sort = m_adoRst.Fields(iCol).Name & " desc, 本所案號 desc "
         m_blnColOrderAsc = False
      Else
         m_adoRst.Sort = m_adoRst.Fields(iCol).Name & " asc, 本所案號 asc "
         m_blnColOrderAsc = True
      End If
      SetRst2Grid
      SetDataListWidth
      grdDataList.Visible = True
    End If
End Sub

'Mark by Lydia 2019/11/04 移到basPublic, 暫時保留
''Added by Lydia 2019/11/01 檢查利益衝突案件之權限(XY特殊權限範圍)
'Private Function PRI_ChkCuFa_Right(ByVal pNo As String, ByVal pSys As String, ByRef outRight As String, ByRef outArea As String) As Boolean
''pNo: 檢查X/Y編號 or 本所案號
''pSys: 檢查系統別: 空白,ALL=>FCP, FG, CFP, PS, CPS
''outRight : 可使用的範圍
''outArea : X/Y編號的管制系統類別
''===============================
''檢查方式:
'' 1.以X/Y編號＋系統類別＋操作員工編號檢查
'' 2.以X/Y編號＋系統類別＋操作者部門檢查
'' 3.檢查是否為該案件(限PA16 is null or <>’1’ )之OA承辦人：需要與Owen確認OA案件性質、A類或C類收文
''前3項其中之一符合即可
''===============================
'Dim strB1 As String
'Dim strChkNo As String 'X/Y編號
'Dim strConB As String
'Dim intJ As Integer
'Dim rsB As New ADODB.Recordset
'
'On Error GoTo ExitProc
'
'    PRI_ChkCuFa_Right = False
'
'    'X/Y編號
'    If InStr("X,Y", Left(pNo, 1)) > 0 Then
'        strChkNo = ChangeCustomerL(pNo)
'    End If
'    outRight = ""
'    outArea = ""
'
'    strConB = "SELECT distinct(CFR02) AS CFRArea FROM CUFA_RIGHT WHERE CFR01='" & Left(strChkNo, 8) & "' "
'    '先取得X/Y編號的管制系統類別
'    intJ = 1
'    Set rsB = ClsLawReadRstMsg(intJ, strConB & " order by 1")
'    If intJ = 1 Then
'         outArea = "" & rsB.GetString(adClipString, , , ",")
'         If Right(outArea, 1) = "," Then outArea = Mid(outArea, 1, Len(outArea) - 1)
'    End If
'    If outArea = "" Then
'         '不管制
'         PRI_ChkCuFa_Right = True
'         GoTo ExitProc
'    End If
'    '+系統類別
'    strConB = strConB & IIf(pSys = "" Or pSys = "ALL", "", " AND CFR02 IN (" & GetAddStr(pSys) & ") ")
'
'' 1.以X/Y編號＋系統類別＋操作員工編號檢查
'    strB1 = strConB & " and cfr03='" & strUserNum & "' "
'    intJ = 1
'    Set rsB = ClsLawReadRstMsg(intJ, strB1 & " order by 1")
'    If intJ = 1 Then
'        outRight = "" & rsB.GetString(adClipString, , , ",")
'    End If
'    If outRight = "" Then
'' 2.以X/Y編號＋系統類別＋操作者部門檢查
'        strB1 = strConB & " and cfr03='" & Pub_StrUserSt03 & "' "
'        intJ = 1
'        Set rsB = ClsLawReadRstMsg(intJ, strB1 & " order by 1")
'        If intJ = 1 Then
'            outRight = "" & rsB.GetString(adClipString, , , ",")
'        End If
'    End If
'
'    If outRight = "" Then
'' 3.檢查是否為該案件(限PA16 is null or <>’1’ )之OA承辦人：需要與Owen確認OA案件性質、A類或C類收文
'         strConB = "select cp01,cp02,cp03,cp04 from caseprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp159=0 and cp14='" & strUserNum & "' "
'         '----專利案的承辦人
'         strB1 = "select '" & strUserNum & "', '" & Me.Name & "', '" & Left(strChkNo, 8) & "', pa01,pa02,pa03,pa04 from patent where " & _
'                      IIf(Left(strChkNo, 1) = "X", " instr(pa26||','||pa27||','||pa28||','||pa29||','||pa30,'" & Left(strChkNo, 8) & "') > 0 ", " substr(pa75,1,8)='" & Left(strChkNo, 8) & "' ") & _
'                     "and nvl(pa16,'0')<>'1' and (pa01,pa02,pa03,pa04) in (" & strConB & ") "
'         '----服務案的承辦人
'         strB1 = strB1 & "union select '" & strUserNum & "', '" & Me.Name & "', '" & Left(strChkNo, 8) & "', sp01,sp02,sp03,sp04 from servicepractice where " & _
'                     IIf(Left(strChkNo, 1) = "X", " instr(sp08||','||sp58||','||sp59||','||sp65||','||sp66,'" & Left(strChkNo, 8) & "') > 0 ", " substr(sp26,1,8)='" & Left(strChkNo, 8) & "' ") & _
'                     "and (sp01,sp02,sp03,sp04) in (" & Replace(strConB, "pa", "sp") & ") "
'         cnnConnection.Execute " insert into R100102_2 (R02201,R02202,R02203,R02204,R02205,R02206,R02207) " & strB1, intJ
'    End If
'
'    If outRight <> "" Then
'       If Right(outRight, 1) = "," Then outRight = Mid(outRight, 1, Len(outRight) - 1)
'       PRI_ChkCuFa_Right = True
'    End If
'
'ExitProc:
'    Set rsB = Nothing
'    If Err.Number <> 0 Then
'        MsgBox Err.Description, vbCritical, msgtext(1110)
'        Resume Next
'    End If
'    Exit Function
'End Function
'
''Added by Lydia 2019/11/01 組合SQL條件 (by各程式)
'Private Function ComConSQL(ByVal iKind As String, ByVal iChkNo As String, ByVal iRight As String, ByVal iArea As String) As String
'Dim ArrCon As Variant
'Dim intR As Integer
'Dim strMidCon  As String
'Dim rsR1 As New ADODB.Recordset
'
'    ComConSQL = ""
'    If iRight = "" And iArea = "" Then Exit Function
'
'    strMidCon = "select R02204||R02205||R02206||R02207 as caseno from R100102_2 " & _
'                       "where R02201='" & strUserNum & "' and R02202='" & Me.Name & "' and R02203='" & Left(iChkNo, 8) & "' group by R02204||R02205||R02206||R02207 "
'    intR = 1
'    Set rsR1 = ClsLawReadRstMsg(intR, strMidCon)
'    If intR = 1 Then
'        strMidCon = GetAddStr(rsR1.GetString(adClipString, , , ","))
'    Else
'        strMidCon = ""
'    End If
'    Set rsR1 = Nothing
'
'    If strMidCon = "" Then
'        ArrCon = Split(iArea, ",")
'        For intR = 0 To UBound(ArrCon)
'           If Trim(ArrCon(intR)) <> "" Then
'              If CheckRight1(ArrCon(intR), iRight) = False Then
'                 Select Case iKind
'                     Case "PA" '專利
'                         If Left(iChkNo, 1) = "X" Then
'                            strMidCon = strMidCon & " OR (instr(PA26||','||PA27||','||PA28||','||PA29||','||PA30, '" & Left(ChangeCustomerL(iChkNo), 8) & "') > 0 and PA01='" & ArrCon(intR) & "')"
'                         ElseIf Left(iChkNo, 1) = "Y" Then
'                            strMidCon = strMidCon & " OR (instr(PA75, '" & Left(ChangeCustomerL(iChkNo), 8) & "') > 0 and PA01='" & ArrCon(intR) & "')"
'                         End If
'                     Case "SP" '服務業務
'                         If Left(iChkNo, 1) = "X" Then
'                            strMidCon = strMidCon & " OR (instr(SP08||','||SP58||','||SP59||','||SP65||','||SP66, '" & Left(ChangeCustomerL(iChkNo), 8) & "') > 0 and SP01='" & ArrCon(intR) & "')"
'                         ElseIf Left(iChkNo, 1) = "Y" Then
'                            strMidCon = strMidCon & " OR (instr(SP26, '" & Left(ChangeCustomerL(iChkNo), 8) & "') > 0 and SP01='" & ArrCon(intR) & "')"
'                         End If
'                 End Select
'              End If
'           End If
'        Next intR
'        If strMidCon <> "" Then strMidCon = " AND NOT(" & Mid(strMidCon, 4) & ") "
'
'    Else  '本所案號
'        If iKind = "PA" Then
'              strMidCon = " AND (PA01||PA02||PA03||PA04) IN (" & strMidCon & ") "
'        ElseIf iKind = "SP" Then
'              strMidCon = " AND (SP01||SP02||SP03||SP04) IN (" & strMidCon & ") "
'        End If
'    End If
'    ComConSQL = strMidCon
'
'End Function
'
''從控制的系統別,比對是否有權限
'Private Function CheckRight1(ByVal pArea As String, ByVal pRights As String) As Boolean
''pArea : 控制的系統別
''pRights: 所有的權限
'Dim arrRight As Variant
'Dim intA As Integer
'
'    CheckRight1 = False
'
'    If pRights = "" Then
'         '全部-無權限
'    Else
'         '有全部權限和部份權限
'         arrRight = Split(pRights, ",")
'         For intA = 0 To UBound(arrRight)
'              If Trim(arrRight(intA)) <> "" And Trim(arrRight(intA)) = pArea Then
'                  CheckRight1 = True '逐一比對,有權限
'                  Exit For
'              End If
'         Next intA
'    End If
'    Exit Function
'
'End Function

'Add by Amy 2023/01/09 改共用函數,並整理程式
Sub StrMenu()
    Dim strAppNo As String, strAppNo1 As String, strContactNo As String
    Dim dblRow As Double 'Add By Sindy 2025/9/3
    
    BolFrom100114 = False
    Me.Enabled = False
    Str01 = ""    '申請人編號
    Str02 = ""    '系統類別
    Str03 = ""    '收文日期(起)
    Str04 = ""    '收文日期(迄)
    Str05 = ""    '案件性質(起)
    Str06 = ""    '案件性質(迄)
    Str07 = ""    '是否含來函資料
    Str01 = Me.Tag
    SetCustData 'Added by Lydia 2023/08/09 預示顯示客戶編號+名稱
    
    '檢查國內外權限
    'Modified by Lydia 2023/08/09 查客戶X編號之案件無權限時，請帶出該客戶曾收文案件之系統類別; ex.杜協理查詢X66247
    'If CheckSR12(Str01) = False Then
    If CheckSR12(Str01, , True) = False Then
        Screen.MousePointer = vbDefault
        Me.Enabled = True
        tmpBol = fnCancelNowFormAndShowParentForm(Me)
        Exit Sub
    End If
    If bolIsL = False Then
        Label1 = "申請人編號："
    Else
        Label1 = "當事人編號："
    End If
    pub_QL05 = pub_QL05 & ";" & Label1 & Str01 & "(案件)" 'Add By Sindy 2025/8/13
    
    lblContact.Caption = ""
    Str02 = IIf(m_Sys <> "ALL", m_Sys, GetAllSysKind(, m_Sys))
    Str03 = m_Date1 '收文日起
    Str04 = IIf(Len(m_Date1) <= 0, "", IIf(Len(m_Date2) <= 0, (ServerDate - 19110000), m_Date2)) '收文日迄
    Str05 = m_Pty1 '案件性質(起)
    Str06 = m_Pty2 '案件性質(迄)
    Str07 = m_CKind '是否含來函資料
    strSQL1 = "" '組字串
    
    '收文
    If m_Type = "1" Then
        If Len(Str03) <> 0 Then
            strSQL1 = strSQL1 + " and CP.cp05>=" & Val(ChangeTStringToWString(Str03))
        End If
        If Len(Str04) <> 0 Then
            strSQL1 = strSQL1 + " and CP.cp05<=" & Val(ChangeTStringToWString(Str04))
        End If
        'Add By Sindy 2025/8/13
        If Len(Str03) <> 0 Or Len(Str04) <> 0 Then
            pub_QL05 = pub_QL05 & ";收文日期：" & Str03 & "-" & Str04
        End If
        '2025/8/13 END
    '配合代理人查詢畫面的條件-發文
    Else
        If Len(Str03) <> 0 Then
            strSQL1 = strSQL1 + " and CP.cp27>=" & Val(ChangeTStringToWString(Str03))
        End If
        If Len(Str04) <> 0 Then
            strSQL1 = strSQL1 + " and CP.cp27<=" & Val(ChangeTStringToWString(Str04))
        End If
        'Add By Sindy 2025/8/13
        If Len(Str03) <> 0 Or Len(Str04) <> 0 Then
            pub_QL05 = pub_QL05 & ";發文日期：" & Str03 & "-" & Str04
        End If
        '2025/8/13 END
    End If
    '案件性質
    If Len(Str05) <> 0 Then
        strSQL1 = strSQL1 + " and CP.cp10>='" & Str05 & "' "
    End If
    If Len(Str06) <> 0 Then
        strSQL1 = strSQL1 + " and CP.cp10<='" & Str06 & "' "
    End If
    'Add By Sindy 2025/8/13
    If Len(Str05) <> 0 Or Len(Str06) <> 0 Then
      pub_QL05 = pub_QL05 & ";案件性質：" & Str05 & "-" & Str06
    End If
    '2025/8/13 END
    '是否含來函資料
    If UCase(Str07) = "N" Then
        strSQL1 = strSQL1 + " and CP.cp09 < 'C' "
        pub_QL05 = pub_QL05 & ";是否含來函資料：不含" 'Add By Sindy 2025/8/13
    End If
    
    '顯示表單上面的值
    Label3.Caption = Me.Tag
       
    '考慮含接洽人編號
    strAppNo1 = Me.Tag 'Modify by Amy 2024/12/10 從下面if 搬出來,共傳入Pub_GetCusCaseSql用
    If Mid(Me.Tag, 10, 1) = "-" Then
        strContactNo = Mid(Me.Tag, 11)
        Me.Tag = Left(Me.Tag, 9)
        strAppNo = Me.Tag
        'strAppNo1 = Left(strAppNo, 8) & "9"
    Else
        Me.Tag = ChangeCustomerL(Me.Tag)
        strAppNo = Me.Tag
        'strAppNo1 = strAppNo
    End If
    If strContactNo <> "" Then
        strSql = "SELECT NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),CU13,ST02,cu111,PCC05" & _
                " FROM CUSTOMER,STAFF,POTCUSTCONT WHERE CU01='" & Left$(Me.Tag, 8) & "' AND CU02='" & Right$(Me.Tag, 1) & "' AND CU13=ST01(+)" & _
                " AND PCC01(+)=CU01 AND PCC02(+)='" & strContactNo & "' "
    Else
        strSql = "SELECT NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),CU13,ST02,cu111,'' PCC05" & _
                " FROM CUSTOMER,STAFF WHERE CU01='" & Left$(Me.Tag, 8) & "' AND CU02='" & Right(Me.Tag, 1) & "' AND CU13=ST01(+) "
    End If
    CheckOC
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
        If IsNull(adoRecordset.Fields(0)) Then
            Label4.Caption = ""
        Else
            Label4.Caption = adoRecordset.Fields(0)
        End If
        If IsNull(adoRecordset.Fields(1)) Then
            Label6.Caption = ""
        Else
            Label6.Caption = adoRecordset.Fields(1)
        End If
        If IsNull(adoRecordset.Fields(2)) Then
            Label7.Caption = ""
        Else
            Label7.Caption = adoRecordset.Fields(2)
        End If
        If CheckStr(adoRecordset.Fields("cu111")) = "Y" Then
            Label3.ForeColor = &HFF&
        Else
            Label3.ForeColor = &H80000012
        End If
       If strContactNo <> "" Then
          lblContactL.Visible = True
          lblContact.Visible = True
          lblContact = "" & adoRecordset.Fields("pcc05")
       End If
    End If
    CheckOC
    'Modify by Amy 2023/10/13 於共用函數判斷取聯絡人編號,原:Pub_GetCusCaseSql(Me.Name, strAppNo, strAppNo1, …)
    'Modify by Amy 2023/01/19 +if bolIsL
    'Modify by Amy 2024/12/10 原Pub_GetCusCaseSql(Me.Name, Me.Tag,...)
    '不是 法務專用
    If bolIsL = False Then
        strSql = Pub_GetCusCaseSql(Me.Name, strAppNo1, m_Sys, bolIsL, ChkPCT.Value, cntLstPayYearSQL, strSQL1, m_Cty1, m_Cty2)
    '法務專用 (前畫面按「法務進度」鈕)
    Else
        strSql = Pub_GetCusCaseSql(Me.Name, strAppNo1, m_Sys, bolIsL, ChkPCT.Value, cntFaSql, strSQL1, m_Cty1, m_Cty2)
    End If
    
    If strContactNo <> "" Then
       strSql = "Select X.* From (" & strSql & ") X Where CNT='" & Left(strAppNo, 8) & strContactNo & "'"
    End If
    'end 2024/12/10
    
    '非法務案
    If bolIsL = False Then
        strSql = strSql & " ORDER BY FSort,本所案號"
        grdDataList.ColWidth(9) = 0
    Else
        strSql = strSql & " ORDER BY FSort,本所案號,收文日"
    End If
    
    CheckOC
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
    If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
        dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/9/3
        If Len(Trim(Str02)) <> 0 Then
            strTemp = Split(Str02, ",")
        End If
        If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
            intCufaCnt = 0
            adoRecordset.MoveFirst
            Do While adoRecordset.EOF = False
                '利益衝突案件：逐案號判斷
                If PUB_ChkCufaByCase(Me.Name, Str02, "" & adoRecordset.Fields("本所案號"), "" & adoRecordset.Fields("cust01") & "," & adoRecordset.Fields("cust02") & "," & adoRecordset.Fields("cust03") & "," & adoRecordset.Fields("cust04") & "," & adoRecordset.Fields("cust05"), "" & adoRecordset.Fields("fcno")) = False Then
                    intCufaCnt = intCufaCnt + 1
                    adoRecordset.Delete
                End If
                adoRecordset.MoveNext
            Loop
            '利益衝突案件：限閱案件
            If intCufaCnt > 0 Then
               pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
               MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
            End If
            If pub_QL04 <> "" Then InsertQueryLog (dblRow) 'Add By Sindy 2025/8/13
            If adoRecordset.RecordCount = 0 Then
                  GoTo JumpToNoData
            End If
        Else
            If pub_QL04 <> "" Then InsertQueryLog (dblRow) 'Add By Sindy 2025/8/13
        End If
        
        Set m_adoRst = adoRecordset.Clone
    Else
        If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/13
JumpToNoData:
        Set m_adoRst = adoRecordset.Clone
        cmdok(0).Enabled = False
        cmdok(1).Enabled = False
        cmdok(2).Enabled = False
        cmdok(3).Enabled = False
        ShowNoData
        Screen.MousePointer = vbDefault
        Me.Enabled = True
        tmpBol = fnCancelNowFormAndShowParentForm(Me)
        Exit Sub
    End If
    grdDataList.FixedCols = 0
    m_adoRst.Sort = "FSort ,本所案號 asc" '先排序,後指定資料集
    SetRst2Grid
    m_blnColOrderAsc = True
    SetDataListWidth
    If bolIsL = False Then
       grdDataList.FixedCols = 4
    Else
       grdDataList.FixedCols = 6
    End If
    Me.Enabled = True
    '有權限者顯示顧問電話諮詢按鈕
    cmdok(8).Enabled = False
    cmdok(8).Visible = False
    If bolIsL = True Then
       cmdok(8).Enabled = IsUserHasRightOfFunction("frm100102_1", strAdd, False)
       cmdok(8).Visible = IsUserHasRightOfFunction("frm100102_1", strAdd, False)
    End If
End Sub

'Added by Lydia 2023/08/09 預示顯示客戶資料
Private Sub SetCustData()
Dim strTmp As String
   If InStr(1, Me.Tag, ",") > 0 Then
      strTmp = Mid(Me.Tag, 1, InStr(1, Me.Tag, ",") - 1)
   Else
      strTmp = Me.Tag
   End If
   Label3.Caption = strTmp
   If Len(Trim(strTmp)) = 9 Then
      strSql = "SELECT NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),CU13,ST02,cu111 FROM CUSTOMER,STAFF WHERE CU01='" & Left$(GetNewFagent(strTmp), 8) & "' AND CU02='" & Right$(GetNewFagent(strTmp), 1) & "' AND CU13=ST01(+)"
   Else
      strSql = "SELECT NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),CU13,ST02,cu111 FROM CUSTOMER,STAFF WHERE CU01='" & Left$(GetNewFagent(strTmp), 8) & "' AND CU02='0' AND CU13=ST01(+) "
   End If
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       If IsNull(adoRecordset.Fields(0)) Then
           Label4.Caption = ""
       Else
           Label4.Caption = adoRecordset.Fields(0)
       End If
       If IsNull(adoRecordset.Fields(1)) Then
           Label6.Caption = ""
       Else
           Label6.Caption = adoRecordset.Fields(1)
       End If
       If IsNull(adoRecordset.Fields(2)) Then
           Label7.Caption = ""
       Else
           Label7.Caption = adoRecordset.Fields(2)
       End If
       If CheckStr(adoRecordset.Fields("cu111")) = "Y" Then
           Label3.ForeColor = &HFF&
       Else
           Label3.ForeColor = &H80000012
       End If
   End If
   CheckOC
End Sub
