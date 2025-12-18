VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090401 
   BorderStyle     =   1  '單線固定
   Caption         =   "撰寫信函"
   ClientHeight    =   6372
   ClientLeft      =   900
   ClientTop       =   1056
   ClientWidth     =   7800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6372
   ScaleWidth      =   7800
   Begin VB.CheckBox ChkINSTdef 
      BackColor       =   &H00FFFFC0&
      Caption         =   "含欄位設定"
      Height          =   225
      Left            =   5460
      TabIndex        =   51
      Top             =   570
      Width           =   1275
   End
   Begin VB.CheckBox ChkINST 
      BackColor       =   &H00FFFFC0&
      Caption         =   "各項指示"
      Height          =   225
      Left            =   4380
      TabIndex        =   50
      Top             =   420
      Width           =   1065
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6840
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdFCMail 
      Caption         =   "FC翻譯案件郵件"
      Enabled         =   0   'False
      Height          =   345
      Index           =   2
      Left            =   40
      TabIndex        =   27
      Top             =   30
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CommandButton cmdFCMail 
      Caption         =   "發FC郵件(承辦署名)"
      Enabled         =   0   'False
      Height          =   345
      Index           =   1
      Left            =   1680
      TabIndex        =   28
      Top             =   30
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton cmdFCMail 
      Caption         =   "發FC郵件(工程師署名)"
      Enabled         =   0   'False
      Height          =   345
      Index           =   0
      Left            =   3450
      TabIndex        =   29
      Top             =   30
      Width           =   1935
   End
   Begin VB.ComboBox Combo8 
      Height          =   300
      Left            =   4131
      Style           =   2  '單純下拉式
      TabIndex        =   6
      Top             =   990
      Width           =   2445
   End
   Begin VB.TextBox txtLetterHead 
      Height          =   270
      Left            =   4545
      MaxLength       =   1
      TabIndex        =   26
      Text            =   "N"
      Top             =   6022
      Width           =   375
   End
   Begin VB.TextBox txtFaxFace 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1620
      MaxLength       =   1
      TabIndex        =   25
      Text            =   "N"
      Top             =   6022
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "回前畫面(&U)"
      Height          =   345
      Left            =   6600
      TabIndex        =   31
      Top             =   30
      Width           =   1100
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Word編輯(&W)"
      Height          =   345
      Left            =   5400
      TabIndex        =   30
      Top             =   30
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      Caption         =   "聯絡人"
      Height          =   1545
      Left            =   72
      TabIndex        =   36
      Top             =   3030
      Width           =   7635
      Begin VB.OptionButton Option5 
         Caption         =   "實體聯絡人名稱"
         Height          =   180
         Left            =   108
         TabIndex        =   20
         Top             =   1200
         Width           =   1575
      End
      Begin VB.OptionButton Option4 
         Caption         =   "聯絡人1名稱"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin MSForms.ComboBox Combo6 
         Height          =   300
         Left            =   1700
         TabIndex        =   21
         Top             =   1110
         Width           =   5865
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "10345;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtConDepJP 
         Height          =   300
         Left            =   1700
         TabIndex        =   19
         Top             =   810
         Width           =   5865
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "10345;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo5 
         Height          =   300
         Left            =   1700
         TabIndex        =   18
         Top             =   480
         Width           =   5865
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "10345;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo4 
         Height          =   300
         Left            =   1700
         TabIndex        =   17
         Top             =   180
         Width           =   5865
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "10345;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人部門(日)"
         Height          =   180
         Left            =   360
         TabIndex        =   46
         Top             =   870
         Width           =   1200
      End
      Begin VB.Label Label5 
         Caption         =   "聯絡人2名稱"
         Height          =   180
         Left            =   360
         TabIndex        =   37
         Top             =   540
         Width           =   1092
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "日文"
      Height          =   180
      Index           =   2
      Left            =   2796
      TabIndex        =   9
      Top             =   1365
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "英文"
      Height          =   180
      Index           =   1
      Left            =   1956
      TabIndex        =   8
      Top             =   1365
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "中文"
      Height          =   180
      Index           =   0
      Left            =   1116
      TabIndex        =   7
      Top             =   1365
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "發信對象："
      Height          =   1410
      Left            =   72
      TabIndex        =   33
      Top             =   1590
      Width           =   7635
      Begin VB.OptionButton Option3 
         Caption         =   "申請人名稱"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   1005
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "FC代理人名稱"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   288
         Width           =   1545
      End
      Begin VB.OptionButton Option6 
         Caption         =   "CF代理人名稱"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   660
         Width           =   1545
      End
      Begin MSForms.ComboBox Combo3 
         Height          =   300
         Left            =   1700
         TabIndex        =   15
         Top             =   960
         Width           =   5865
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "10345;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo7 
         Height          =   300
         Left            =   1700
         TabIndex        =   13
         Top             =   600
         Width           =   5865
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "10345;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Left            =   1700
         TabIndex        =   11
         Top             =   270
         Width           =   5865
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "10345;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   2775
      MaxLength       =   2
      TabIndex        =   4
      Top             =   375
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   2535
      MaxLength       =   1
      TabIndex        =   3
      Top             =   375
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1575
      MaxLength       =   6
      TabIndex        =   2
      Top             =   375
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1116
      MaxLength       =   3
      TabIndex        =   1
      Top             =   375
      Width           =   495
   End
   Begin MSForms.TextBox Text6 
      Height          =   450
      Index           =   1
      Left            =   1230
      TabIndex        =   24
      Top             =   5520
      Width           =   6465
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "11404;794"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text6 
      Height          =   450
      Index           =   0
      Left            =   1230
      TabIndex        =   23
      Top             =   5070
      Width           =   6465
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "11404;794"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   450
      Left            =   1230
      TabIndex        =   22
      Top             =   4620
      Width           =   6465
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "11404;794"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1110
      TabIndex        =   5
      Top             =   660
      Width           =   6600
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "11642;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblSendMailDt 
      AutoSize        =   -1  'True
      Caption         =   "寄件日期:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   5550
      TabIndex        =   49
      Top             =   420
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label14 
      Caption         =   "中文信函請用開窗信紙!!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   3915
      TabIndex        =   48
      Top             =   1320
      Width           =   2580
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "信函格式："
      Height          =   180
      Index           =   1
      Left            =   3195
      TabIndex        =   47
      Top             =   1050
      Width           =   900
   End
   Begin VB.Label Label4 
      Caption         =   "是否印信頭：        （N:不印）"
      Height          =   225
      Index           =   1
      Left            =   3420
      TabIndex        =   45
      Top             =   6045
      Width           =   2370
   End
   Begin VB.Label Label4 
      Caption         =   "是否印傳真封面：        （N:不印）"
      Height          =   225
      Index           =   4
      Left            =   105
      TabIndex        =   44
      Top             =   6045
      Width           =   2685
   End
   Begin VB.Label lblClose 
      Caption         =   "lblClose"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3200
      TabIndex        =   43
      Top             =   420
      Width           =   975
   End
   Begin VB.Label Label12 
      Height          =   180
      Left            =   1905
      TabIndex        =   42
      Top             =   1065
      Width           =   975
   End
   Begin VB.Label Label11 
      Height          =   180
      Left            =   1110
      TabIndex        =   41
      Top             =   1065
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "申請人備註："
      Height          =   180
      Left            =   105
      TabIndex        =   40
      Top             =   5655
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "代理人備註："
      Height          =   180
      Left            =   105
      TabIndex        =   39
      Top             =   5205
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "案件備註："
      Height          =   180
      Left            =   105
      TabIndex        =   38
      Top             =   4755
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "信函語文："
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   35
      Top             =   1350
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "申請國家："
      Height          =   180
      Left            =   90
      TabIndex        =   34
      Top             =   1065
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "案件名稱："
      Height          =   180
      Left            =   90
      TabIndex        =   32
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   180
      Left            =   90
      TabIndex        =   0
      Top             =   420
      Width           =   975
   End
End
Attribute VB_Name = "frm090401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/7/27 日文已改抓Table
'Memo by Lydia 2021/09/23 改成Form2.0 ; Combo1~Combo7、txtConDepJP、Text5~Text7
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
'2010/9/10 配合CFP修改日文格式
'Modified by Morgan 2021/8/12 智慧財產法院-->智慧財產及商業法院
Option Explicit

'edit by nickc 2006/07/12
'Dim pA(T_PA) As String, i As Integer
Dim pa() As String, i As Integer
Dim intWhere As Integer
Dim cu(0 To 30) As String
Dim fa(0 To 30) As String
Dim cfa(0 To 30) As String
Dim cp(0 To 5) As String
Dim np(0 To 5) As String
Dim ptm(0 To 3) As String
Dim na(0 To 1) As String
Dim strTemp As String
'Add By Sindy 2017/3/7
Dim m_Y52269FA16Mail As String '巨京代表信箱
Dim strCP14 As String
Dim strCP14ST17 As String
'2017/3/7 END
'Add by Morgan 2010/9/7
Dim m_strCP45 As String 'CF彼所案號
Dim m_strCP09 As String
'end 2010/9/7
Dim m_strCP44 As String 'CF代理人
Dim m_strCP44_FA16 As String 'Added by Morgan 2024/1/15 CF代理人代表信箱
Dim m_strFCAgent As String 'FC代理人
'Modified by Lydia 2021/02/08 將申請人1~5設為陣列,並且全部取代m_strCustomer=> m_CustNo(1)
'Dim m_strCustomer As String '申請人
Dim m_CustNo(1 To 5) As String '申請人編號1~5
Dim m_CustName(1 To 5) As String '申請人名稱1~5
Dim m_CustMemo(1 To 5) As String '申請人備註1~5
'end 2021/02/08
Dim m_strContact1(2) As String '聯絡人1
Dim m_strContact2(2) As String '聯絡人2
Dim m_strContact3(2) As String '實體聯絡人
Dim m_strFax(0 To 3) As String '傳真封面資料 0=FAX1,1=TEL1,2=FAX2,3=TEL2
Dim bolRetry As Boolean '是否已發生錯誤且重試
Dim bolFA As Boolean '是否可選代理人
Dim m_strConDepJp As String  '聯絡人部門(日)
Dim m_iLang As Integer '預設定稿語文
Dim m_bEMail As Boolean '是否以EMail通知
'2008/5/1 add by sonia
Dim custtype As String   '個人/公司
Dim CASENAME As String   '案件名稱
Dim nationname As String '申請國家
Dim casetype1 As String  '專利商標種類
Dim casetype2 As String  '案件別
Dim casetype3 As String  '卷宗性質
Dim casetype4 As String  '客戶案件案號
Dim casetype5 As String  '分所案號　　　　　2008/9/18 ADD BY SONIA
Dim caselaw As String    '來函條款　　　　　2008/9/19 ADD BY SONIA
Dim casepaper As String  '爭議案來函文書    2008/9/24 ADD BY SONIA
Dim CaseNo As String     '本所案號
Dim custarea As String   '業務區
Dim custsales As String  '智權人員
Dim stCP12 As String, stCP13 As String
'2008/5/1 end
Dim AppNo As String      '申請案號   2008/6/25 add by sonia
'Add by Morgan 2008/7/17
Dim m_Zip As String '郵遞區號
Dim m_Contact As String '接洽人
Dim m_CU80 As String '客戶狀態
Dim m_Dept As String   '2006/12/11 ADD BY SONIA
Dim m_CU104 As String 'Add by Morgan 2008/8/6
Dim m_Combo8 As String   '信函格式   2008/9/16  ADD BY SONIA
Dim m_QCase As String 'Add By Sindy 2013/2/18 記錄目前所查詢的案號資料
Dim m_bolReadOK As Boolean 'Added by Morgan 2013/3/8
Dim m_CP115 As String 'Add By Sindy 2013/4/2
Dim m_CompNo As String 'Added by Morgan 2014/1/17 特殊公司別
Dim strTemplatePath As String 'Add By Sindy 2014/8/15
'Dim bolIsECase As Boolean 'Add By Sindy 2014/8/26 是否為E化案件
Dim m_StrUserST03 As String 'Add By Sindy 2014/9/12 操作者部門
Public EMailType As String 'Add By Sindy 2014/9/15 郵件格式
Public strAttach As String 'Add By Sindy 2014/10/7 Mail要加附件
Public bolFrom1105Callme As Boolean 'Add By Sindy 2014/10/8 判斷是否有frm1105定稿維護呼叫此作業
Public OutCallCP10 As String 'Add By Sindy 2015/1/8 傳入案件性質
Public OutCallProcCP10 As String 'Add By Sindy 2019/5/23 傳入要帶入主旨的案件性質
Public OutCallCP09 As String 'Add By Sindy 2015/1/8 傳入總收文號
'Add By Sindy 2015/5/29
Public m_ET01 As String '定稿別
Public m_ET03 As String '處理狀況
'2015/5/29 END
Public m_ET99 As String '份數'Add By Sindy 2017/10/6
Dim m_AboutDeadLine As String 'Added by Morgan 2015/5/14
Dim m_CP49 As String    '2009/5/26 add by sonia
'Added by Lydia 2015/11/02
Public mPcnt1 As Integer
Public mPcnt2 As Integer
Dim m2FileName As String
'Added by Lydia 2015/11/11
Public strLoadPath As String '讀取附件路徑
Private Const TSMailName As String = "工作通知單" 'Added by Lydia 2017/06/16 FC翻譯案件郵件寄給對方的檔案名稱
Dim strMemoX As String, strMemoY As String, strMemoCase As String 'Added by Lydia 2017/08/03 客戶,代理人,案件的各項備註內容
Dim strKeyX As String, strKeyY As String, strKeyCase As String 'Added by Lydia 2020/06/04 有各項備註的客戶,代理人,案件編號
Dim strSavePath1 As String 'Added by Lydia 2018/06/25 FC翻譯郵件的附件存放路徑
Dim m_ChildCP44 As String 'Added by Lydia 2019/02/26 EPC之子案指定代理人(CFP)
Dim bolDateType As Boolean   'add by sonia 2019/11/19 True:西元年月日,False:民國年月日
Dim m_strIT10 As String 'Added by Lydia 2020/06/12 各項指示之使用部門
'Added by Lydia 2020/07/16 法律所案源收文
Dim stLos04 As String
Dim stLos04Name As String '案源之介紹人名稱
Public m_IDSCP09 As String 'Added by Morgan 2020/12/30 IDS收文號
Dim bIsBPFCase As Boolean, stNP09 As String, stNP23 As String 'Added by Morgan 2021/2/3 寶齡富錦通知信(選英文,FC代理人)
Dim m_TMGoods As String   'add by sonia 2021/4/20 多類別之商品
Public strReturnSheet As String 'Added by Morgan 2023/6/21 回覆單
'Add By Sindy 2025/8/12
Dim m_T727RecvNo As String 'T台灣案分析總收文號
Dim m_T727CP43No As String 'T台灣案分析相關總收文號
'2025/8/12 END


'讀取案件資料
Private Function CaseNoCheck() As Boolean
   Dim Cancel As Boolean
   
   'Modify By Sindy 2013/2/18 增加檢查是否已有執行查詢動作
   'If Combo1.ListCount = 0 Then
   If Combo1.ListCount = 0 Or (m_QCase <> Trim(Text1) & Trim(Text2) & Trim(Text3) & Trim(Text4)) Then
   '2013/2/18 End
      Text3_Validate Cancel
      Text4_LostFocus
   End If
   
   'Modified by Morgan 2013/3/8
   'CaseNoCheck = True
   CaseNoCheck = m_bolReadOK
End Function

'Modify By Sindy 2024/7/23 mark,改用共用函數 PUB_GetLen
''Modify by Morgan 2007/3/29 加所有申請人選項
'Private Function GetCustName(strCaseNo As String, strLang As String, Optional ByVal bolAll As Boolean = False, Optional ByVal PreStr As String, Optional ByVal bolNum As Boolean = False) As String
'On Error GoTo ErrHnd
'
'   If bolAll = True Then
'      'Modify By Sindy 2011/2/18 增加LC43,LC44,LC45,LC46,HC24,HC25,HC26,HC27
'      strSql = "select pa26,pa27,pa28,pa29,pa30 from patent where " & ChgPatent(strCaseNo)
'      strSql = strSql & " Union Select tm23,tm78,tm79,tm80,tm81 From Trademark Where " & ChgTradeMark(strCaseNo)
'      strSql = strSql & " Union Select LC11,LC43,LC44,LC45,LC46 From Lawcase Where " & ChgLawcase(strCaseNo)
'      strSql = strSql & " Union Select HC11,HC24,HC25,HC26,HC27 From Hirecase Where " & ChgHirecase(strCaseNo)
'      strSql = strSql & " Union Select sp08,sp58,sp59,sp65,sp66 From ServicePractice Where " & ChgService(strCaseNo)
'
'      strSql = "select C1.CU04, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90, C1.CU06" & _
'         ",C2.CU04, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90, C2.CU06" & _
'         ",C3.CU04, C3.CU05||' '||C3.CU88||' '||C3.CU89||' '||C3.CU90, C3.CU06" & _
'         ",C4.CU04, C4.CU05||' '||C4.CU88||' '||C4.CU89||' '||C4.CU90, C4.CU06" & _
'         ",C5.CU04, C5.CU05||' '||C5.CU88||' '||C5.CU89||' '||C5.CU90, C5.CU06" & _
'         " from (" & strSql & ") X,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5" & _
'         " where C1.CU01(+)=substr(PA26,1,8) And C1.CU02(+)=substr(PA26,9,1)" & _
'         " and C2.CU01(+)=substr(PA27,1,8) And C2.CU02(+)=substr(PA27,9,1)" & _
'         " and C3.CU01(+)=substr(PA28,1,8) And C3.CU02(+)=substr(PA28,9,1)" & _
'         " and C4.CU01(+)=substr(PA29,1,8) And C4.CU02(+)=substr(PA29,9,1)" & _
'         " and C5.CU01(+)=substr(PA30,1,8) And C5.CU02(+)=substr(PA30,9,1)"
'   Else
'      strSql = "Select CU04, CU05||' '||CU88||' '||CU89||' '||CU90, CU06 From Customer, Patent Where substr(PA26,1,8)=CU01 And substr(PA26,9,1)=CU02 And " & ChgPatent(strCaseNo)
'      strSql = strSql & " Union Select CU04, CU05||' '||CU88||' '||CU89||' '||CU90, CU06 From Customer, Trademark Where substr(TM23,1,8)=CU01 And substr(TM23,9,1)=CU02 And " & ChgTradeMark(strCaseNo)
'      strSql = strSql & " Union Select CU04, CU05||' '||CU88||' '||CU89||' '||CU90, CU06 From Customer, Lawcase Where substr(LC11,1,8)=CU01 And substr(LC11,9,1)=CU02 And " & ChgLawcase(strCaseNo)
'      strSql = strSql & " Union Select CU04, CU05||' '||CU88||' '||CU89||' '||CU90, CU06 From Customer, Hirecase Where substr(HC05,1,8)=CU01 And substr(HC11,9,1)=CU02 And " & ChgHirecase(strCaseNo)
'      strSql = strSql & " Union Select CU04, CU05||' '||CU88||' '||CU89||' '||CU90, CU06 From Customer, ServicePractice Where substr(SP08,1,8)=CU01 And substr(SP08,9,1)=CU02 And " & ChgService(strCaseNo)
'   End If
'   CheckOC
'   With adoRecordset
'      .CursorLocation = adUseClient
'      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'      If .RecordCount > 0 Then
'         Select Case strLang
'            Case "1" '中文
'               GetCustName = "" & .Fields(0).Value
'               If bolAll = True Then
'                  For intI = 1 To 4
'                     strExc(1) = Trim("" & .Fields(3 * intI).Value)
'                     If strExc(1) <> "" Then
'                        GetCustName = GetCustName & vbCrLf & PreStr & strExc(1)
'                     End If
'                  Next
'               End If
'            Case "2" '英文
'               'Modify by Morgan 2007/7/25 加控制斷行
'               'GetCustName = "" & .Fields(1).Value
'               'If bolAll = True Then
'               '   For intI = 1 To 4
'               '      strExc(1) = Trim("" & .Fields(1 + 3 * intI).Value)
'               '      If strExc(1) <> "" Then
'               '         GetCustName = GetCustName & vbCrLf & PreStr & strExc(1)
'               '      End If
'               '   Next
'               'End If
'               If Not IsNull(.Fields(1)) Then
'                  If bolAll = False Then
'                     GetCustName = SplitTitle(.Fields(1), PUB_GetLen(PreStr))
'                  Else
'                     If bolNum = True And Trim("" & .Fields(1 + 3 * 1)) <> "" Then
'                        GetCustName = SplitTitle("1." & .Fields(1), PUB_GetLen(PreStr))
'                     Else
'                        GetCustName = SplitTitle(.Fields(1), PUB_GetLen(PreStr))
'                     End If
'                     For intI = 1 To 4
'                        strExc(1) = Trim("" & .Fields(1 + 3 * intI).Value)
'                        If strExc(1) <> "" Then
'                           If bolNum = True Then
'                              GetCustName = GetCustName & vbCrLf & PreStr & SplitTitle(intI + 1 & "." & strExc(1), PUB_GetLen(PreStr) + 1)
'                           Else
'                              GetCustName = GetCustName & vbCrLf & PreStr & SplitTitle(strExc(1), PUB_GetLen(PreStr))
'                           End If
'                        End If
'                     Next
'                  End If
'               End If
'               'end 2007/7/25
'
'            Case "3" '日文
'               GetCustName = "" & .Fields(2).Value
'               If bolAll = True Then
'                  For intI = 1 To 4
'                     strExc(1) = Trim("" & .Fields(2 + 3 * intI).Value)
'                     If strExc(1) <> "" Then
'                        GetCustName = GetCustName & vbCrLf & PreStr & strExc(1)
'                     End If
'                  Next
'               End If
'         End Select
'      End If
'   End With
'
'ErrHnd:
'   If Err.Number <> 0 Then MsgBox Err.Description
'End Function

'Modify By Sindy 2014/9/18 Move basQuery
'Private Function GetEngPatKindName(sPTM01 As String, sPTM02 As String) As String
'On Error GoTo ErrHnd
'
'   strSql = "select PTM05 from patenttrademarkmap where PTM01='" & sPTM01 & "' AND PTM02='" & sPTM02 & "'"
'   CheckOC
'   With adoRecordset
'      .CursorLocation = adUseClient
'      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'      '若有資料
'      If .RecordCount > 0 Then
'         GetEngPatKindName = " " & .Fields(0)
'         If GetEngPatKindName = " Patent" Then GetEngPatKindName = ""
'      End If
'   End With
'
'ErrHnd:
'   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
'End Function

Private Sub Command1_Click()
'Debug.Print Format(Now, "nn:ss:") & Right(Format(Timer, ".00"), 2) & "-->Begin"
Dim strTemp1
Dim strTemp2
Dim ii As Integer
Dim jj As Integer
Dim ss As Integer
Dim stReceiver As String
Dim bolChinaFormat As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim m_CP43 As String    '2008/9/19 ADD BY SONIA
Dim bolNoCopy As Boolean 'Added by Morgan 2017/3/14

   bIsBPFCase = False: stNP09 = "": stNP23 = "" 'Added by Morgan 2021/2/2
   
   '檢查本所案號
   If Text1.Text = "" Or Text2.Text = "" Then
       MsgBox "請輸入本所案號!!!", vbExclamation + vbOKOnly
       Text1.SetFocus
       Text1_GotFocus
       Exit Sub
   End If

   If CaseNoCheck = False Then Exit Sub 'Add by Morgan 2004/10/4
   
   Call GetStrCustomer 'Added by Lydia 2020/09/11 抓申請人1的編號
      
   'Add By Sindy 2025/8/12
   m_T727RecvNo = "" 'T台灣案分析總收文號
   m_T727CP43No = "" 'T台灣案分析相關總收文號
   If Text1 = "T" And pa(10) = "000" Then
      strExc(1) = ""
      If Combo8.ItemData(Combo8.ListIndex) <> 0 Then
         strExc(1) = PUB_Num2DocNo(Combo8.ItemData(Combo8.ListIndex))
      End If
      If strExc(1) <> "" Then
         strExc(0) = "Select CP43,CP09,CP24,CP27,CP49 FROM CaseProgress" & _
                     " WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10='727' AND CP09='" & strExc(1) & "'" & _
                     " AND CP27 IS NULL AND CP57 IS NULL"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            m_T727RecvNo = RsTemp.Fields("CP09")
            m_T727CP43No = Trim("" & RsTemp.Fields("CP43"))
         End If
      End If
   End If
   '2025/8/12 END

   'Added by Lydia 2023/10/04 FMP案待客戶最終指示相關控管
   If Pub_StrUserSt03 = "F21" And PUB_ChkFMP970mail("3", Text1, Text2, Text3, Text4, strSql) = True Then
      If strSql <> "" Then
         MsgBox strSql, vbInformation
      End If
   End If
   'end 2023/10/04
   '2008/9/16 ADD BY SONIA
   m_Combo8 = "00"
   Select Case Combo8.Text
      Case "核駁　　　　　　1002"    '核駁
         Select Case Text1.Text
            Case "CFP"
               m_Combo8 = "11"
            '2008/9/26 add by sonia
            Case "T"
               If pa(10) = "000" Then
                  'ADD BY SONIA 2015/9/16 台灣案提申請意見書後之申請或分割核駁
                  StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='202' AND NVL(CP27,0)>0 AND CP09<'B' "
                  rsA.CursorLocation = adUseClient
                  rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                  If rsA.RecordCount > 0 Then
                     m_Combo8 = "56"
                  End If
                  If rsA.State <> adStateClosed Then rsA.Close
                  Set rsA = Nothing
               Else
                  m_Combo8 = "51"
               End If
            'Add By Sindy 2009/07/02
            Case "TF"
               If pa(10) = "000" Then
               Else
                  m_Combo8 = "53"
               End If
            '2009/07/02 End
            '2008/9/26 end
         End Select
      Case "最終核駁　　　　1006"    '最終核駁
         If Text1.Text = "CFP" Then
            '2011/3/16 modify by sonia
            'm_Combo8 = "12"
            If pa(9) = "101" Then
               m_Combo8 = "12"
            Else
               m_Combo8 = "11"
            End If
            '2011/3/16 end
         End If
      Case "通知要求選取　　1206"    '通知要求選取
         If Text1.Text = "CFP" Then
            m_Combo8 = "13"
         End If
      Case "檢索報告　　　　1209"    '檢索報告
         '2008/12/1 MODIFY BY SONIA 加P案函知客戶檢索報告
         'If Text1.Text = "CFP" Then
         If Text1.Text = "CFP" Or Text1.Text = "P" Then
            m_Combo8 = "14"
         End If
         '2008/12/1 END
      '2009/3/3 ADD BY SONIA 國際初步審查報告1216
      Case "國際初步審查報告1216"    '檢索報告
         If Text1.Text = "P" And pa(9) = "056" Then
            m_Combo8 = "15"
         End If
      '2009/3/3 END
      Case "核駁前先行通知　1202"    '核駁前先行通知
         If Text1.Text = "T" Then
            If pa(10) = "000" Then   '台灣依來函條款再分類
               StrSQLa = "Select CP49 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                         IIf(m_T727CP43No <> "", " AND CP09='" & m_T727CP43No & "' ", " AND CP10='1202' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' ")
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  m_CP49 = "" & rsA.Fields("CP49")
                  m_Combo8 = "00"
'MOVE BY SONIA 2015/9/15 移至GetLaw,改為共用(台灣核駁也要用)
                  GetLaw   '依條款代碼取得條款名稱caselaw
'                  caselawTemp = Split(m_CP49, ",")
'                  caselawcount = UBound(caselawTemp) + 1
'                  Dim strLawItem As String
'                  strLawItem = ""
'                  For i = 0 To UBound(caselawTemp)
'                     If strLawItem <> "" Then strLawItem = strLawItem & ","
'                     If Len(caselawTemp(i)) = 1 Then
'                        strLawItem = strLawItem & caselawTemp(i)
'                     Else
'                        strLawItem = strLawItem & Mid(caselawTemp(i), 2, 1)
'                     End If
'                  Next i
''                  '2008/11/19 add by sonia
''                  '第23條第1項第1款(AXX)及第23條第1項第2款(BXX)及第23條第1項第11款(FXX)
''                  If caselawcount = 3 And InStr(rsA.Fields("CP49"), "A") > 0 And InStr(rsA.Fields("CP49"), "B") > 0 And InStr(rsA.Fields("CP49"), "F") > 0 Then
''                     m_Combo8 = "2C"
''                     caselaw = "第5條第2項及第23條第1項第1款、第2款及第11款"
''                  '2008/11/19 end
''                  '2010/10/7 add by sonia
''                  '第23條第1項第1款(AXX)及第23條第1項第2款(BXX)及第23條第1項第13款(HXX)
''                  ElseIf caselawcount = 3 And InStr(rsA.Fields("CP49"), "A") > 0 And InStr(rsA.Fields("CP49"), "B") > 0 And InStr(rsA.Fields("CP49"), "H") > 0 Then
''                     m_Combo8 = "2D"
''                     caselaw = "第5條第2項及第23條第1項第1款、第2款及第13款"
''                  '2010/10/7 end
''                  '2011/2/25 Add By Sindy
''                  '第23條第1項第1款(AXX)及第13款(HXX)
''                  ElseIf caselawcount = 2 And InStr(rsA.Fields("CP49"), "A") > 0 And InStr(rsA.Fields("CP49"), "H") > 0 Then
''                     m_Combo8 = "2E"
''                     caselaw = "第5條第2項、第23條第1項第1款及第13款"
''                  '2011/2/25 End
''                  '第23條第1項第1款(AXX)及第23條第1項第2款(BXX)
''                  ElseIf caselawcount = 2 And InStr(rsA.Fields("CP49"), "A") > 0 And InStr(rsA.Fields("CP49"), "B") > 0 Then
''                     m_Combo8 = "21"
''                     caselaw = "第5條第2項、第23條第1項第1款及第2款"
''                  '第23條第1項第2款(BXX)及第23條第1項第11款(FXX)
''                  ElseIf caselawcount = 2 And InStr(rsA.Fields("CP49"), "B") > 0 And InStr(rsA.Fields("CP49"), "F") > 0 Then
''                     m_Combo8 = "22"
''                     caselaw = "第23條第1項第2款及第11款"
''                  '第23條第1項第2款(BXX)及第23條第1項第13款(HXX)
''                  ElseIf caselawcount = 2 And InStr(rsA.Fields("CP49"), "B") > 0 And InStr(rsA.Fields("CP49"), "H") > 0 Then
''                     m_Combo8 = "23"
''                     caselaw = "第23條第1項第2款及第13款"
''                  '第23條第1項第11款(FXX)及第23條第1項第13款(HXX)
''                  ElseIf caselawcount = 2 And InStr(rsA.Fields("CP49"), "F") > 0 And InStr(rsA.Fields("CP49"), "H") > 0 Then
''                     m_Combo8 = "24"
''                     caselaw = "第23條第1項第11款及第13款"
''                  '第23條第1項第1款(AXX)
''                  ElseIf caselawcount = 1 And InStr(rsA.Fields("CP49"), "A") > 0 Then
''                     m_Combo8 = "25"
''                     caselaw = "第5條第2項及第23條第1項第1款"
''                  '第23條第1項第2款(BXX)
''                  ElseIf caselawcount = 1 And InStr(rsA.Fields("CP49"), "B") > 0 Then
''                     m_Combo8 = "26"
''                     caselaw = "第23條第1項第2款"
''                  '第23條第1項第11款(FXX)
''                  ElseIf caselawcount = 1 And InStr(rsA.Fields("CP49"), "F") > 0 Then
''                     m_Combo8 = "27"
''                     caselaw = "第23條第1項第11款"
''                  '第23條第1項第12款(GXX)
''                  ElseIf caselawcount = 1 And InStr(rsA.Fields("CP49"), "G") > 0 Then
''                     m_Combo8 = "28"
''                     caselaw = "第23條第1項第12款"
''                  '第23條第1項第13款(HXX)
''                  ElseIf caselawcount = 1 And InStr(rsA.Fields("CP49"), "H") > 0 Then
''                     m_Combo8 = "29"
''                     caselaw = "第23條第1項第13款"
''                  '第23條第1項第14款(IXX)
''                  ElseIf caselawcount = 1 And InStr(rsA.Fields("CP49"), "I") > 0 Then
''                     m_Combo8 = "2A"
''                     caselaw = "第23條第1項第14款"
''                  '第23條第1項第15款(JXX)
''                  ElseIf caselawcount = 1 And InStr(rsA.Fields("CP49"), "J") > 0 Then
''                     m_Combo8 = "2B"
''                     caselaw = "第23條第1項第15款"
''                  '其他條款(AXX)
''                  Else
''                     m_Combo8 = "2Z"
''                     caselaw = "第XX條第X項第X款"
''                  End If
'                  'Modify By Sindy 2012/7/11 101年7月1日商標修法
'                  If caselawcount = 3 Then
'                     '第23條第1項第1款(AXX)及第23條第1項第2款(BXX)及第23條第1項第11款(FXX)
'                     If InStr(strLawItem, "A") > 0 And InStr(strLawItem, "B") > 0 And InStr(strLawItem, "F") > 0 Then
'                        m_Combo8 = "2C"
'                        caselaw = "第5條第2項及第23條第1項第1款、第2款及第11款"
'                     '第23條第1項第1款(AXX)及第23條第1項第2款(BXX)及第23條第1項第13款(HXX)
'                     ElseIf InStr(strLawItem, "A") > 0 And InStr(strLawItem, "B") > 0 And InStr(strLawItem, "H") > 0 Then
'                        m_Combo8 = "2D"
'                        caselaw = "第5條第2項及第23條第1項第1款、第2款及第13款"
'                     End If
'                  ElseIf caselawcount = 2 Then
'                     'Modify By Sindy 2012/9/6
'                     '第29條第1項第1款(AXX)及第29條第1項第3款(BXX)
'                     If InStr(strLawItem, "A") > 0 And InStr(strLawItem, "B") > 0 Then
'                        m_Combo8 = "21"
'                        caselaw = "第18條第2項及第29條第1項第1款及第3款"
'                     '2012/9/6 End
'                     'Add By Sindy 2012/9/6
'                     '第29條第1項第2款(AXX)及第29條第1項第3款(CXX)
'                     ElseIf InStr(strLawItem, "A") > 0 And InStr(strLawItem, "C") > 0 Then
'                        m_Combo8 = "2M"
'                        caselaw = "第18條第2項、第29條第1項第2款及第3款"
'                     '2012/9/6 End
'                     '第29條第1項第3款(AXX)及第30條第1項第10款(HXX)
'                     ElseIf InStr(strLawItem, "A") > 0 And InStr(strLawItem, "H") > 0 Then
'                        m_Combo8 = "2E"
'                        caselaw = "第18條第2項、第29條第1項第2款及第30條第1項第10款"
'                     '第29條第1項第1款(BXX)及第30條第1項第8款(FXX)
'                     ElseIf InStr(strLawItem, "B") > 0 And InStr(strLawItem, "F") > 0 Then
'                        m_Combo8 = "22"
'                        caselaw = "第18條第2項、第29條第1項第1款及第30條第1項第8款"
'                     '第29條第1項第1款(BXX)及第30條第1項第10款(HXX)
'                     ElseIf InStr(strLawItem, "B") > 0 And InStr(strLawItem, "H") > 0 Then
'                        m_Combo8 = "23"
'                        caselaw = "第18條第2項、第29條第1項第1款及第30條第1項第10款"
'                     '第30條第1項第8款(FXX)及第30條第1項第10款(HXX)
'                     ElseIf InStr(strLawItem, "F") > 0 And InStr(strLawItem, "H") > 0 Then
'                        m_Combo8 = "24"
'                        caselaw = "第30條第1項第8款及第10款"
'                     '第29條第1項第2款(CXX)及第30條第1項第8款(FXX)
'                     ElseIf InStr(strLawItem, "C") > 0 And InStr(strLawItem, "F") > 0 Then
'                        m_Combo8 = "2G"
'                        caselaw = "第18條第2項、第29條第1項第2款及第30條第1項第8款"
'                     '第29條第1項第3款(AXX)及第30條第1項第8款(FXX)
'                     ElseIf InStr(strLawItem, "A") > 0 And InStr(strLawItem, "F") > 0 Then
'                        m_Combo8 = "2H"
'                        caselaw = "第18條第2項、第29條第1項第3款及第30條第1項第8款"
'                     '第29條第3項(MXX)及第30條第1項第8款(FXX)
'                     ElseIf InStr(strLawItem, "M") > 0 And InStr(strLawItem, "F") > 0 Then
'                        m_Combo8 = "2J"
'                        caselaw = "第18條第2項、第29條第3項及第30條第1項第8款"
'                     '第29條第3項(MXX)及第30條第1項第10款(HXX)
'                     ElseIf InStr(strLawItem, "M") > 0 And InStr(strLawItem, "H") > 0 Then
'                        m_Combo8 = "2K"
'                        caselaw = "第18條第2項、第29條第3項及第30條第1項第10款"
'                     'Add By Sindy 2012/7/17
'                     '第29條第3項(MXX)及第30條第1項第11款(GXX)
'                     ElseIf InStr(strLawItem, "M") > 0 And InStr(strLawItem, "G") > 0 Then
'                        m_Combo8 = "2L"
'                        caselaw = "第18條第2項、第29條第3項及第30條第1項11款"
'                     '2012/7/17 End
'                     End If
'                  ElseIf caselawcount = 1 Then
'                     '第29條第1項第3款(AXX)
'                     If InStr(strLawItem, "A") > 0 Then
'                        m_Combo8 = "25"
'                        caselaw = "第18條第2項及第29條第1項第3款"
'                     '第29條第1項第1款(BXX)
'                     ElseIf InStr(strLawItem, "B") > 0 Then
'                        m_Combo8 = "26"
'                        caselaw = "第18條第2項及第29條第1項第1款"
'                     '第29條第1項第2款(CXX)
'                     ElseIf InStr(strLawItem, "C") > 0 Then
'                        m_Combo8 = "2F"
'                        caselaw = "第18條第2項及第29條第1項第2款"
'                     '第30條第1項第8款(FXX)
'                     ElseIf InStr(strLawItem, "F") > 0 Then
'                        m_Combo8 = "27"
'                        caselaw = "第30條第1項第8款"
'                     '第30條第1項第11款(GXX)
'                     ElseIf InStr(strLawItem, "G") > 0 Then
'                        m_Combo8 = "28"
'                        caselaw = "第30條第1項11款"
'                     '第30條第1項第10款(HXX)
'                     ElseIf InStr(strLawItem, "H") > 0 Then
'                        m_Combo8 = "29"
'                        caselaw = "第30條第1項第10款"
'                     '第30條第1項第12款(IXX)
'                     ElseIf InStr(strLawItem, "I") > 0 Then
'                        m_Combo8 = "2A"
'                        caselaw = "第30條第1項第12款"
'                     '第30條第1項第13款(JXX)
'                     ElseIf InStr(strLawItem, "J") > 0 Then
'                        m_Combo8 = "2B"
'                        caselaw = "第30條第1項第13款"
'                     '第29條第3項(MXX)
'                     ElseIf InStr(strLawItem, "M") > 0 Then
'                        m_Combo8 = "2I"
'                        caselaw = "第18條第2項及第29條第3項"
'                     End If
'                  End If
'                  If m_Combo8 = "00" Then
'                     '其他條款(AXX)
'                     m_Combo8 = "2Z"
'                     caselaw = "第XX條第X項第X款"
'                  End If
'                  '2012/7/11 End
'END 2015/9/16
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            End If
         End If
      Case "勝訴　　　　　　1003", "撤銷原處分－勝　1402"    '勝訴,撤銷原處分-勝
         If Text1.Text = "T" Then  '依相關總收文號案件性質再分類
            '2008/10/17 加入撤銷原處分
            If Combo8.Text = "勝訴　　　　　　1003" Then
               StrSQLa = "Select C2.CP10 FROM CaseProgress C1,CaseProgress C2 WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                         IIf(m_T727CP43No <> "", " AND C1.CP09='" & m_T727CP43No & "' ", " AND C1.CP10='1003' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C' ") & " AND C1.CP43=C2.CP09(+) "
            Else
               StrSQLa = "Select C2.CP10 FROM CaseProgress C1,CaseProgress C2 WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                         IIf(m_T727CP43No <> "", " AND C1.CP09='" & m_T727CP43No & "' AND C1.CP24='1' ", " AND C1.CP10='1402' AND C1.CP24='1' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C' ") & " AND C1.CP43=C2.CP09(+) "
            End If
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Select Case rsA.Fields(0)
                  'modify by sonia 2019/5/27 +623部分廢止
                  Case "601", "603", "605", "623"   '異議,評定,廢止
                     Select Case pa(10)
                        Case "000"            '台灣
                           m_Combo8 = "31"
                        Case Else             '非台灣
                           m_Combo8 = "32"
                     End Select
                  'modify by sonia 2019/5/27 +1620被部分廢止（理由）
                  Case "1602", "1604", "1606", "1620" '被異議（理由）,被評定（理由）,被廢止（理由）
                     Select Case pa(10)
                        Case "000"            '台灣
                           m_Combo8 = "33"
                        Case Else             '非台灣
                           m_Combo8 = "34"
                     End Select
                  'modify by sonia 2019/5/27 +624部分廢止答辯
                  Case "602", "604", "606", "624"   '異議答辯,評定答辯,廢止答辯
                     Select Case pa(10)
                        Case "000"            '台灣
                           m_Combo8 = "35"
                        Case Else             '非台灣
                           '2017/6/14 modify by sonia 台->大廢止(撤銷)答辯獨立出來T-147816
                           'm_Combo8 = "36"
                           'modify by sonia 2019/5/27 +624部分廢止答辯
                           If rsA.Fields(0) = "606" Or rsA.Fields(0) = "624" Then
                              m_Combo8 = "3E"
                           Else
                              m_Combo8 = "36"
                           End If
                           '2017/6/14 end
                     End Select
                  Case "401"                  '訴願
                     Select Case pa(10)
                        Case "000"            '台灣
                           If pa(28) = "1" Then     '申請
                              m_Combo8 = "37"
                           Else
                              m_Combo8 = "39"       '爭議
                           End If
                        Case Else             '非台灣
                           m_Combo8 = "38"
                     End Select
                  Case "403"                  '行政訴訟
                     Select Case pa(10)
                        Case "000"            '台灣
                           If pa(28) <> "1" Then    '爭議
                              m_Combo8 = "3B"
                           End If
                     End Select
                  Case "406"                  '參加訴願
                     Select Case pa(10)
                        Case "000"            '台灣
                           If pa(28) <> "1" Then    '爭議
                              m_Combo8 = "3C"
                           End If
                     End Select
                  Case "410"                  '行政上訴答辯
                     Select Case pa(10)
                        Case "000"            '台灣
                           '2009/3/3 MODIFY BY SONIA 取消卷宗性質控制T-065430
                           'If pa(28) <> "1" Then    '爭議
                              m_Combo8 = "3D"
                           'End If
                     End Select
               End Select
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
         End If
      'modify by sonia 2019/11/19 +部分勝部分敗1006(T-217896)
      Case "敗訴　　　　　　1004", "撤銷原處分－敗　1402", "部分勝部分敗　　1006"   '敗訴,撤銷原處分-敗,部分勝部分敗
         If Text1.Text = "T" Then  '依相關總收文號案件性質再分類
            '2008/10/17 加入撤銷原處分
            'modify by sonia 2019/11/19 +部分勝部分敗　　1006
            If Combo8.Text = "敗訴　　　　　　1004" Or Combo8.Text = "部分勝部分敗　　1006" Then
               StrSQLa = "Select C2.CP10 FROM CaseProgress C1,CaseProgress C2 WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                         IIf(m_T727CP43No <> "", " AND C1.CP09='" & m_T727CP43No & "' ", " AND C1.CP10 in ('1004','1006') AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C'") & " AND C1.CP43=C2.CP09(+) "
            Else
               StrSQLa = "Select C2.CP10 FROM CaseProgress C1,CaseProgress C2 WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                         IIf(m_T727CP43No <> "", " AND C1.CP09='" & m_T727CP43No & "' AND C1.CP24='2' ", " AND C1.CP10='1402' AND C1.CP24='2' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C'") & " AND C1.CP43=C2.CP09(+) "
            End If
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Select Case rsA.Fields(0)
                  'modify by sonia 2019/5/27 +623部分廢止
                  Case "601", "603", "605", "623"   '異議,評定,廢止
                     Select Case pa(10)
                        Case "000"            '台灣
                           m_Combo8 = "41"
                        Case Else             '非台灣
                           m_Combo8 = "42"
                           '2015/7/15 add by sonia 大陸異議敗訴無期限,改新格式 T-183743
                           If pa(10) = "020" And rsA.Fields(0) = "601" Then m_Combo8 = "42A"
                     End Select
                  'modify by sonia 2019/5/27 +1620被部分廢止（理由）
                  Case "1602", "1604", "1606", "1620" '被異議（理由）,被評定（理由）,被廢止（理由）
                     Select Case pa(10)
                        Case "000"            '台灣
                           m_Combo8 = "43"
                        Case Else             '非台灣
                           m_Combo8 = "44"
                     End Select
                  'modify by sonia 2019/5/27 +624部分廢止答辯
                  Case "602", "604", "606", "624"   '異議答辯,評定答辯,廢止答辯
                     Select Case pa(10)
                        Case "000"            '台灣
                           m_Combo8 = "45"
                        Case Else             '非台灣
                           m_Combo8 = "46"
                     End Select
                  Case "401"                  '訴願
                     Select Case pa(10)
                        Case "000"            '台灣
                           m_Combo8 = "47"
                        Case Else             '非台灣
                           m_Combo8 = "48"
                     End Select
                  Case "403", "407"           '行政訴訟,參加訴訟
                     Select Case pa(10)
                        Case "000"            '台灣
                           '2012/5/1 cancel by sonia 林副理說被申請核駁案件定稿同爭議案件
                           'If pa(28) <> "1" Then    '爭議
                              m_Combo8 = "4B"
                           'End If
                     End Select
                  Case "1404"                 '通知參加訴願(未參加)
                     Select Case pa(10)
                        Case "000"            '台灣
                           m_Combo8 = "4C"
                     End Select
                  'Add By Sindy 2012/4/23
                  Case "408"                 '行政訴訟上訴
                     Select Case pa(10)
                        Case "000"            '台灣
                           m_Combo8 = "4D"
                     End Select
                  '2012/4/23 End
               End Select
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
         End If
      '2008/9/24 ADD BY SONIA
      Case "部分核駁　　　　1205"    '部分核駁
         If Text1.Text = "T" Then
            If pa(10) <> "000" Then   '非台灣
               m_Combo8 = "52"
            End If
         End If
      '2008/9/24 END
      'Add By Sindy 2009/07/21
      '2010/11/5 MODIFY BY SONIA 加台->大通知修正1702 T-171559
      Case "審查報告　　　　1201", "通知修正　　　　1702"    '審查報告,通知修正
         If Text1.Text = "T" Then
            If pa(10) = "000" Then   '台->台
               m_Combo8 = "54"
            Else '台->大
               m_Combo8 = "55"
            End If
         End If
      '2009/07/21 End
      '2008/10/8 ADD BY SONIA
      Case "通知準備程序　　1203", "通知言詞辯論　　1204"   '通知準備程序,通知言詞辯論
         If Text1.Text = "T" Then
            If pa(10) = "000" Then   '台灣依來函所選程序是否為勝訴再分類
               StrSQLa = "Select C2.CP24 FROM CaseProgress C1,CaseProgress C2 WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                         IIf(m_T727CP43No <> "", " AND C1.CP09='" & m_T727CP43No & "' ", " AND C1.CP10 IN ('1203','1204') AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C'") & " AND C1.CP43=C2.CP09(+) "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               m_Combo8 = "62"        '2009/6/15 ADD BY SONIA T-065429
               If rsA.RecordCount > 0 Then
                  If Not IsNull(rsA.Fields(0)) And rsA.Fields(0) = "1" Then
                     m_Combo8 = "61"
'2009/6/15 CANCEL BY SONIA T-065429
'                  Else
'                     m_Combo8 = "62"
                  End If
               End If
            End If
         End If
      Case "通知行政上訴答辯1406"    '通知行政上訴答辯
         If Text1.Text = "T" Then
            If pa(10) = "000" Then   '台灣
               m_Combo8 = "63"
            End If
         End If
      'modify by sonia 2019/5/27 +1619被部分廢止
      Case "被異議　　　　　1601", "被評定　　　　　1603", "被廢止　　　　　1605", "被部分廢止　　　1619"  '被異議,被評定,被廢止
         If Text1.Text = "T" Then
            If pa(10) = "000" Then   '台灣
               m_Combo8 = "64"
            End If
         End If
      'modify by sonia 2019/5/27 +1620被部分廢止（理由）
      Case "被異議（理由）　1602", "被評定（理由）　1604", "被廢止（理由）　1606", "被部分廢止（理由）1620"  '被異議（理由）,被評定（理由）,被廢止（理由）
         If Text1.Text = "T" Then
            If pa(10) = "000" Then   '台灣
               m_Combo8 = "65"
            Else '非台灣
               'Add By Sindy 2009/08/12
               'modify by sonia 2019/5/27 +1620被部分廢止（理由）
               If Combo8.Text = "被廢止（理由）　1606" Or Combo8.Text = "被部分廢止（理由）1620" Then
                  m_Combo8 = "84"
               Else
                  m_Combo8 = "83"
               End If
            End If
         'Modify By Sindy 2009/08/21
         ElseIf Text1.Text = "TF" Then
            m_Combo8 = "85"
         '2009/08/21 End
         End If
      Case "對方補充理由　　1609", "發回補答辯　　　1613"   '對方補充理由,2009/2/26加1613發回補答辯
         If Text1.Text = "T" Then
            If pa(10) = "000" Then   '台灣
               m_Combo8 = "66"
            End If
         End If
      Case "對方答辯　　　　1618", "補證據　　　　　1617"   '對方答辯,2008/10/20加1617補證據
         If Text1.Text = "T" Then
            If pa(10) = "000" Then   '台灣 台->台
               m_Combo8 = "67"
            'Add By Sindy 2012/4/23
            Else '台->大
               m_Combo8 = "6A"
            End If
         End If
      Case "智慧局答辯函　　1709"    '智慧局答辯函
         If Text1.Text = "T" Then
            If pa(10) = "000" Then   '台灣
               m_Combo8 = "68"
            End If
         End If
      '2008/10/8 END
      
'      'Add By Sindy 2009/10/26
'      Case "變更申請案號　　1718"    '變更申請案號
'         If Text1.Text = "T" Then
'            If pa(10) = "000" Then   '台灣
'               m_Combo8 = "69"
'            End If
'         End If
'      '2009/10/26 End
      
      'Add By Sindy 2009/08/12
      Case "通知復審答辯　　1404"
         If Text1.Text = "T" Then
            StrSQLa = "Select C2.CP10,C1.CP43 FROM CaseProgress C1,CaseProgress C2 WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                         IIf(m_T727CP43No <> "", " AND C1.CP09='" & m_T727CP43No & "' ", " AND C1.CP10='1404' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C'") & " AND C1.CP43=C2.CP09(+) "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               If Left(Trim(rsA.Fields(1)), 1) = "C" Then '未答辯
                  Select Case rsA.Fields(0)
                     'modify by sonia 2019/5/27 +1620被部分廢止（理由）
                     Case "1602", "1604", "1606", "1620"   '被異議,被評定,被廢止
                        Select Case pa(10)
                           Case "000"            '台灣
                           Case Else             '非台灣
                              m_Combo8 = "81"
                        End Select
                  End Select
               ElseIf Left(Trim(rsA.Fields(1)), 1) = "A" Or Left(Trim(rsA.Fields(1)), 1) = "B" Then '已答辯
                  Select Case rsA.Fields(0)
                     'modify by sonia 2019/5/27 +624部分廢止答辯
                     Case "602", "604", "606", "624" '異議答辯,評定答辯,廢止答辯
                        Select Case pa(10)
                           Case "000"            '台灣
                           Case Else             '非台灣
                              m_Combo8 = "82"
                        End Select
                  End Select
               End If
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
         End If
      '2009/08/12 End
      '2012/4/27 add by sonia 自程序定稿移過來
      Case "對方撤回　　　　1610"    '對方撤回
         If Text1.Text = "T" Then
            If pa(10) = "000" Then   '台灣
               m_Combo8 = "6B"
            End If
         End If
      '2013/1/21 ADD BY SONIA CFP分析941帶核駁格式
      Case "分析　　　　　　941"    '分析
         Select Case Text1.Text
            Case "CFP"
               m_Combo8 = "11"
            'Add By Sindy 2013/3/29
            Case "P"
               m_Combo8 = "16"
            '2013/3/29 End
         End Select
      '2013/1/21 END
      
      'Added by Morgan 2016/3/18
      'Modified by Morgan 2022/8/24 +部分准駁1009
      Case "分析　　　　　　1001", "分析　　　　　　1002", "分析　　　　　　1009"
         If Text1.Text = "P" Then m_Combo8 = "16"
      
      'ADD BY SONIA 2014/6/4 P-105534
      Case "專利權評價報告  1209"    '專利權評價報告
         If Text1.Text = "P" Then
            m_Combo8 = "17"
         End If
      'END 2014/6/4
      
   End Select
   '2008/9/16 END
   
   '2006/12/11 ADD BY SONIA 開放FF案件之權限
   m_Dept = GetStaffDepartment(strUserNum)
   'Modify by Morgan 2007/1/25 Systemkind_g
   'strTemp1 = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
   strTemp1 = Split(Replace(UCase(Systemkind_g), ",,", ""), ",")
   'end 2007/1/25
   strTemp2 = Split(Replace(UCase(Text1.Text), ",,", ""), ",")
   For ii = 0 To UBound(strTemp2)
       ss = 0
       For jj = 0 To UBound(strTemp1)
           If strTemp2(ii) = strTemp1(jj) Then
               ss = 1
               Exit For
           End If
       Next jj
       If ss = 0 Then
          Select Case m_Dept
             'Modify by Morgan 2007/4/11 加F61
             'Modify by Morgan 2008/4/8 加F81
             Case "F21", "F23", "F61", "F81"  '開放F21,F23使用P,PS,CFP,CPS權限
                If Text1.Text = "P" Or Text1.Text = "PS" Or Text1.Text = "CFP" Or Text1.Text = "CPS" Then
                   Exit For
                End If
             Case "F10", "F11"    '開放F10,F11使用T權限
                'modify by sonia 2014/4/28
                'If Text1.Text = "T" Then
                If InStr(Text1.Text, "T") > 0 Then
                   Exit For
                End If
          End Select
          '2006/12/11 END
          '2010/3/25 ADD BY SONIA 檢查跨部門權限
          If CheckSR09(strUserNum, Text1, "Y", False, Text1, Text2, Text3, Text4) = True Then
             Exit For
          End If
          '2010/3/25 END
           ss = MsgBox(strUserName & " 沒有 " & strTemp2(ii) & " 的權限!! ", , "USER 權限問題")
           Text1.SetFocus
           Text1_GotFocus
           Exit Sub
       End If
   Next ii
   If Option2.Value = False And Option3.Value = False And Option6.Value = False Then
       MsgBox "請選擇FC代理人或CF代理人或申請人名稱!!!", vbCritical
       Exit Sub
   End If
   If Option1(0).Value = False And Option1(1).Value = False And Option1(2).Value = False Then
       MsgBox "請選擇信函語文種類!!!", vbCritical
       Exit Sub
   End If
   'add by sonia 2021/10/14
   If Option2.Value = True And Combo2 = "" Then
       MsgBox "FC代理人名稱空白!!!", vbCritical
       Exit Sub
   End If
   If Option6.Value = True And Combo7 = "" Then
      'Added by Morgan 2021/11/17 無CF代理人改提醒可繼續，因IDS會與新案同步作業，另P案的代理人撰稿也會在新案未發文的狀況下寫信--玫音,紹仁
      'MsgBox "CF代理人名稱空白!!!", vbCritical
      'Exit Sub
      If MsgBox("CF代理人名稱空白，是否確定要繼續？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
      'end 2021/11/17
   End If
   If Option3.Value = True And Combo3 = "" Then
       MsgBox "申請人名稱空白!!!", vbCritical
       Exit Sub
   End If
   'end 2021/10/14

   Screen.MousePointer = vbHourglass
   '2008/5/1 add by sonia 專利處要求加印
   '個人/公司
   'Modify By Sindy 2012/5/24 DECODE(CU15,'0','台端','貴公司')==>DECODE(CU15,'0','台端','1','貴公司','貴單位')
   strExc(0) = "SELECT DECODE(CU15,'0','台端','1','貴公司','貴單位') FROM CUSTOMER WHERE " & ChgCustomer(m_CustNo(1))
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      custtype = "" & RsTemp.Fields(0).Value
   End If
   '申請國家
   nationname = Label12.Caption
   If nationname = "台灣" Then nationname = ""    '台灣案不印
   '案件名稱casename,專利商標種類casetype1,案件別casetype2,卷宗性質casetype3,casepaper爭議案文書
   casetype1 = "": casetype2 = "": casetype3 = "": casepaper = ""
   Select Case Text1.Text
      Case "P", "CFP", "FCP"
         'Modify By Sindy 2013/3/13 雅娟:有關通知客戶的定稿,若有英文名稱,則定稿中帶出來的專利名稱應用(),以與中文名稱作區別
         'CASENAME = pa(5) & pa(6) & pa(7)
         CASENAME = Trim(pa(5)) & IIf(pa(6) <> "", "(" & Trim(pa(6)) & ")", "") & IIf(pa(7) <> "", "（" & Trim(pa(7)) & "）", "")
         '2013/3/13 End
         If Label11 <> "000" And Text1 = "P" Then
            strExc(0) = "SELECT decode(pa09,'013',decode(pa08,'1','標準專利','2','短期專利',PTM04),PTM04) FROM PATENTTRADEMARKMAP,PATENT WHERE PA01='" & Text1 & "' AND PA02='" & Text2 & "' AND PA03='" & Text3 & "' AND PA04='" & Text4 & "'" & _
                        " AND PTM01='1' AND PTM02=PA08"
         Else
            strExc(0) = "SELECT PTM03 FROM PATENTTRADEMARKMAP,PATENT WHERE PA01='" & Text1 & "' AND PA02='" & Text2 & "' AND PA03='" & Text3 & "' AND PA04='" & Text4 & "'" & _
                        " AND PTM01='1' AND PTM02=PA08"
         End If
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            casetype1 = "" & RsTemp.Fields(0).Value
         End If
         casetype2 = "專利"
         Select Case pa(23)
            Case "1"
               casetype3 = "申請"
            Case "2"
               casetype3 = "異議"
            Case "3"
               casetype3 = "舉發"
         End Select
      Case "T", "TF", "CFT", "FCT"
         CASENAME = IIf(pa(131) <> "", pa(131), pa(5)) & pa(6) & pa(7) 'Modify By Sindy 2015/7/1 +TM131
         If Label11 <> "000" Then
            strExc(0) = "SELECT DECODE(TM08,'1','商標','2','商標','3','商標','4','服務商標','5','服務商標','6','服務商標',PTM04) " & _
                        " FROM TRADEMARK, PatentTrademarkMap WHERE TM01='" & Text1 & "' AND TM02='" & Text2 & "' AND TM03='" & Text3 & "' AND TM04='" & Text4 & "' And TM08=PTM02(+) And '2'=PTM01 "
         Else
            strExc(0) = "SELECT PTM03||Decode(TM58, Null, '', Decode(instr(TM58, '原為聯合商標'), 0, Decode(instr(TM58, '原為服務標章'), 0, Decode(instr(TM58, '原為聯合服務標章'), 0, '', '（原為聯合服務標章）'), '（原為服務標章）'), '（原為聯合商標）') ) " & _
                        " FROM TRADEMARK, PatentTrademarkMap WHERE TM01='" & Text1 & "' AND TM02='" & Text2 & "' AND TM03='" & Text3 & "' AND TM04='" & Text4 & "' And TM08=PTM02(+) And '2'=PTM01 "
         End If
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            casetype1 = "" & RsTemp.Fields(0).Value
         End If
         'casetype2 = "商標"   '因商標種類已為商標
         Select Case pa(28)
            Case "1"
               casetype2 = "核駁"          '2008/9/19 ADD BY SONIA
               If pa(22) <> "" Then casetype2 = "註冊" '2010/6/11 ADD BY SONIA有專用期 T-159946
               If Left(m_Combo8, 1) = "0" Or Left(m_Combo8, 1) = "2" Then
                  casetype3 = "註冊申請"
                  '2009/8/19 add by sonia 已發註冊證為「註冊」，未發證為「申請」
                  If Text1.Text = "CFT" Then
                     If pa(21) = "" Then
                        casetype3 = "申請"
                     Else
                        casetype3 = "註冊"
                     End If
                  End If
                  '2009/8/19 end
               Else
                  casetype3 = "註冊"
               End If
               '2012/5/1 ADD BY SONIA 卷宗性質為申請者再以案件性質判斷casetype3(T-175718)
               'modify by sonia 2017/9/1 +判斷Combo8 T-176979同時有1604,1606
               'StrSQLa = "Select MAX(CP05||CP10) FROM CaseProgress WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
                         " AND CP10 IN ('101','602','1602','604','1604','606','1606','1601','1603','1605') AND INSTR('" & Combo8 & "',CP10)>0 "
               'modify by sonia 2019/5/27 +624部分廢止答辯,+1619被部分廢止,1620被部分廢止（理由）
               StrSQLa = "Select MAX(CP05||CP10) FROM CaseProgress WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
                         " AND CP10 IN ('101','602','1602','604','1604','606','1606','1601','1603','1605','624','1619','1620') "
               'modify by sonia 2019/5/27 +624部分廢止答辯,1620被部分廢止（理由）
               If InStr(Combo8, "602") > 0 Or InStr(Combo8, "604") > 0 Or InStr(Combo8, "606") > 0 Or InStr(Combo8, "624") > 0 _
               Or InStr(Combo8, "1602") > 0 Or InStr(Combo8, "1604") > 0 Or InStr(Combo8, "1606") > 0 Or InStr(Combo8, "1620") > 0 Then
                  StrSQLa = StrSQLa & " AND INSTR('" & Combo8 & "',CP10)>0 "
               End If
               'end 2017/9/1
               If rsA.State <> adStateClosed Then rsA.Close
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If IsNull(rsA.Fields(0)) = False Then
                  Select Case Mid(rsA.Fields(0), 9)
                     Case "602", "1602", "1601"
                        casetype3 = "異議"
                     Case "604", "1604", "1603"
                        If pa(10) = "000" Then
                           casetype3 = "評定"
                        Else
                           casetype3 = "裁定"
                        End If
                     'modify by sonia 2019/5/27 +624部分廢止答辯,+1619被部分廢止,+1620被部分廢止（理由）
                     Case "606", "1606", "1605", "624", "1619", "1620"
                        If pa(10) = "000" Then
                           casetype3 = "廢止"
                        Else
                           casetype3 = "撤銷"
                        End If
                  End Select
               Else
                  casetype3 = "　　"    '中間接進來則留空白
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               '2012/5/1 END
            Case "2"
               casetype2 = "註冊"          '2008/9/19 ADD BY SONIA
               casetype3 = "異議"
               If pa(10) = "000" Then
                  casepaper = "審定書"
               Else
                  '2015/7/15 MODIFY BY SONIA T-183743
                  'casepaper = "裁定書"
                  casepaper = "決定書"
               End If
            Case "3"
               casetype2 = "註冊"          '2008/9/19 ADD BY SONIA
               If pa(10) = "000" Then
                  casetype3 = "評定"
                  casepaper = "書"
               Else
                  casetype3 = "裁定"
                  '2009/11/5 MODIFY BY SONIA T-133711
                  'casepaper = "裁定書"
                  casepaper = "書"
               End If
            Case "4"
               casetype2 = "註冊"          '2008/9/19 ADD BY SONIA
               If pa(10) = "000" Then
                  casetype3 = "廢止"
                  casepaper = "處分書"
               Else
                  casetype3 = "撤銷"
                  casepaper = "裁定書"
               End If
         End Select
      'Modify By Sindy 2009/07/24 增加LIN系統類別'modify by sonia 2019/7/30 +ACS系統類別
      'modify by sonia 2019/7/30 +ACS系統類別
      Case "L", "CFL", "FCL", "LIN", "ACS"
         casetype2 = "法務"
      Case "LA"
         casetype2 = "顧問"
      Case "CFC"
         CASENAME = pa(5) & pa(6) & pa(7)
         casetype2 = "美國著作權"
      Case "TB"
         CASENAME = pa(5) & pa(6) & pa(7)
         casetype2 = "條碼"
      Case "TC"
         CASENAME = pa(5) & pa(6) & pa(7)
         casetype2 = "著作權"
      Case "TD"
         CASENAME = pa(5) & pa(6) & pa(7)
         casetype2 = "網域名稱"
      Case "TM"
         CASENAME = pa(5) & pa(6) & pa(7)
         casetype2 = "監視系統"
      Case "TR"
         CASENAME = pa(5) & pa(6) & pa(7)
         casetype2 = "商業司查詢"
   End Select
   '本所案號
   CaseNo = Text1 & "-" & Text2
   If Text3 <> "0" Or Text4 <> "00" Then
      CaseNo = CaseNo & "-" & Text3 & "-" & Text4
   End If
   
   'Added by Lydia 2020/07/16 法律所案源收文：撰寫信函落款帶最新A類承辦，同時加帶聯絡人(案源介紹人)
   If strSrvDate(1) >= 法律所案源收文啟用日 And Me.Text1.Text <> "" And InStr(Me.Text1.Text, "L") > 0 Then
        stCP13 = GetCaseCP14(Me.Text1.Text, Me.Text2.Text, Me.Text3.Text, Me.Text4.Text, stLos04, stLos04Name)
        stCP12 = GetSalesArea(stCP13)
   Else
   'end 2020/07/16
        '業務區及智權人員
        stCP13 = PUB_GetAKindSalesNo(Me.Text1.Text, Me.Text2.Text, Me.Text3.Text, Me.Text4.Text)
        stCP12 = GetSalesArea(stCP13)
   End If 'Added by Lydia 2020/07/16
   
   'Modified by Morgan 2015/5/20
   'custarea = GetDepartmentName(stCP12)
   'custsales = GetStaffName(stCP13)
   custarea = PUB_GetLetterSalesZone(stCP13)
   If custarea = "" Then
      'Added by Morgan 2020/3/3
      '專利國內部(P1)要加職稱
      If Left(stCP12, 2) = "P1" Then
         custarea = "專利國內部" & Replace(GetStaffST20(stCP13), "代", "", 1, 1)
      Else
      'end 2020/3/3
         custarea = GetDepartmentName(stCP12)
      End If 'Added by Morgan 2020/3/3
   End If
   custsales = PUB_GetLetterSalesName(stCP13)
   If custsales = "" Then
      custsales = GetStaffName(stCP13)
   End If
   'end 2015/5/20
   '2008/5/1 end
   
   'Removed by Morgan 2015/5/20
   ''Added by Morgan 2015/4/16
   ''A4004張詠翔的落款加帶葉招進副理
   'If stCP13 = "A4004" Then
   '   custsales = "葉招進副理、" & custsales
   ''A4012江維宬的落款加帶賴皇瑞副理
   'ElseIf stCP13 = "A4012" Then
   '   custsales = "賴皇瑞副理、" & custsales
   'End If
   ''end 2015/4/16
   
   '2011/8/5 add by sonia 68096中三杜副總的客戶定稿特別控制(定稿交邱素蓮)
   '該客戶所有案件最後收文智權人員在職則不印業務區而智權人員改為中三杜副總（ＸＸＸ）
   '                          離職則正常列印
   If stCP13 = "68096" Then
      strExc(0) = "select st02 from staff,(select max(cp05||cp13) cp13 from ( " & _
                  "      Select cp05,cp13 From patent, CaseProgress Where pa26='" & GetPrjPeopleNum1(Me.Text1.Text & "-" & Me.Text2.Text & "-" & Me.Text3.Text & "-" & Me.Text4.Text) & "' and pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04 and cp09<'B' " & _
                  "union Select cp05,cp13 From trademark, CaseProgress Where tm23='" & GetPrjPeopleNum1(Me.Text1.Text & "-" & Me.Text2.Text & "-" & Me.Text3.Text & "-" & Me.Text4.Text) & "' and tm01=cp01 and tm02=cp02 and tm03=cp03 and tm04=cp04 and cp09<'B' " & _
                  "union Select cp05,cp13 From lawcase, CaseProgress Where lc11='" & GetPrjPeopleNum1(Me.Text1.Text & "-" & Me.Text2.Text & "-" & Me.Text3.Text & "-" & Me.Text4.Text) & "' and lc01=cp01 and lc02=cp02 and lc03=cp03 and lc04=cp04 and cp09<'B' " & _
                  "union Select cp05,cp13 From servicepractice, CaseProgress Where sp08='" & GetPrjPeopleNum1(Me.Text1.Text & "-" & Me.Text2.Text & "-" & Me.Text3.Text & "-" & Me.Text4.Text) & "' and sp01=cp01 and sp02=cp02 and sp03=cp03 and sp04=cp04 and cp09<'B' " & _
                  "union Select cp05,cp13 From hirecase, CaseProgress Where hc05='" & GetPrjPeopleNum1(Me.Text1.Text & "-" & Me.Text2.Text & "-" & Me.Text3.Text & "-" & Me.Text4.Text) & "' and hc01=cp01 and hc02=cp02 and hc03=cp03 and hc04=cp04 and cp09<'B' " & _
                  ")) aa where substr(aa.cp13,9)=st01(+) and st04='1'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         custarea = ""
         custsales = custsales & "（" & RsTemp.Fields(0).Value & "）"
      End If
   End If
   '2011/8/5 end
   
   '信函語文--中文
   If Option1(0).Value Then
      'Modify by Morgan 2008/7/22 申請人只用台灣格式--郭
      '申請人
      If Option3.Value = True Then
         'Add by Morgan 2008/8/6
         If m_CU104 <> "" Then
            MsgBox "此客戶之收件人名稱與客戶名稱不同，請勿自行修改收件人名稱！", vbExclamation
         End If
         stReceiver = m_CustNo(1)
         
         'Removed by Morgan 2016/11/29 取消改發文時提醒程序(E化案件由系統自動產生)
         ''Added by Morgan 2016/11/10
         'Modified by Morgan 2017/3/14 改操作人員為案件的智權人員時彈訊息詢問是否
         bolNoCopy = False
         If stCP13 = strUserNum Then
            If PUB_GetCC(pa(1), pa(2), pa(3), pa(4), strExc(1), strExc(2)) Then
               If MsgBox("請注意，本案有設副本收件人是否也要通知？" & vbCrLf & vbCrLf & "(選""是""將會另外開啟副本收件人的Word)", vbExclamation + vbYesNo + vbDefaultButton1) = vbYes Then
                  WordChinese True, strExc(1), strExc(2)
               Else
                  bolNoCopy = True
               End If
            End If
         End If
         ''end 2016/11/10
         'end 2016/11/29
                  
         WordChinese , , , bolNoCopy

      Else
         'Modify by Morgan 2006/10/2 加大陸格式
         'WordChinese
         stReceiver = ""
         bolChinaFormat = False
         '2008/10/29 CANCEL BY SONIA
         'If Text1.Text = "P" Or Text1.Text = "PS" Or Text1.Text = "CFP" Or Text1.Text = "CPS" Then
         '2008/10/29 END
            'FC代理人
            If Option2.Value = True Then
               stReceiver = m_strFCAgent
               bolChinaFormat = CheckIsChina(stReceiver)
            'CF代理人
            ElseIf Option6.Value = True Then
               stReceiver = m_strCP44
            End If
            
            'Added by Morgan 2021/12/13 CF代理人應該只用大陸格式 Ex:P-128287 (無理人時) --紹仁
            If Option6.Value = True Then
               bolChinaFormat = True
            Else
            'end 2021/12/13
               bolChinaFormat = CheckIsChina(stReceiver)
            End If
         'End If
         If bolChinaFormat Then
            WordChinese1
         Else
            WordChinese
         End If
         'end 2006/10/2
      End If
   '信函語文--英文
   ElseIf Option1(1).Value Then
   
      'Added by Morgan 2021/2/3 寶齡富錦 Y55435 特殊控制
      If pa(75) = "Y55435" And Option2.Value = True Then
         bIsBPFCase = True
         strExc(0) = "SELECT CP10,NP08,NP09,NP23 FROM caseprogress,NEXTPROGRESS" & _
            " where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & IIf(OutCallCP09 <> "", " and cp09='" & OutCallCP09 & "'", " and cp09>'C'") & _
            " and cp27||cp57 is null and np01(+)=cp09 and NP06 IS NULL" & strNpSqlOfNoSalesDuty & _
            " order by np09 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            stNP09 = "" & RsTemp.Fields("np09")
            '約定期限
            If pa(1) = "P" Then
               '所限-14日
               stNP23 = CompDate(2, -14, RsTemp.Fields("np08"))
            Else
               stNP23 = "" & RsTemp.Fields("np23")
            End If
         End If
      End If
      'end 2021/2/3
      
      WordEnglish
   '信函語文--日文
   Else
      WordJapan
   End If
   Screen.MousePointer = vbDefault
'Debug.Print Format(Now, "nn:ss:") & Right(Format(Timer, ".00"), 2) & "-->End"
End Sub

'Add by Morgan 2006/10/2
'抓代理人/申請人國籍
Private Function CheckIsChina(p_Code As String) As Boolean
   Dim strCode As String
   strCode = Left(p_Code & "000", 9)
   strExc(0) = "select FA10 from fagent where fa01='" & Left(strCode, 8) & "' and fa02='" & Right(strCode, 1) & "'" & _
      " union all select CU10 from customer where cu01='" & Left(strCode, 8) & "' and cu02='" & Right(strCode, 1) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      '2008/7/23 modify by sonia
      'If "" & RsTemp.Fields(0) = "020" Or "" & RsTemp.Fields(0) = "013" Or "" & RsTemp.Fields(0) = "044" Then
      If "" & RsTemp.Fields(0) <> "000" Then
         CheckIsChina = True
      End If
   End If
End Function

Private Sub Command2_Click()
    Unload Me
End Sub

'Add By Sindy 2017/3/7
Private Sub TradeMarkP29Mail()
Dim adoRst As ADODB.Recordset
Dim objOutLook As Object
Dim objMail As Object
Dim strTM05 As String, strTM09 As String
Dim strContent As String, strCuName As String
Dim strApp1Addr As String, strCaseNo As String, strGoods As String
Dim strP20Mgr As String
Dim pbolDone As Boolean, strCP09 As String, strCP10 As String
Dim intMaxEEP02 As Integer
Dim strTM44 As String, strTM45 As String 'Add By Sindy 2025/2/21
   
   '查詢商標資料
   'Modify By Sindy 2025/2/21 +,tm44,tm45
   strExc(0) = "select tm05,tm09,tm23,nvl(tm24,nvl(tm25,tm26)) App1Addr,nvl(cu04,decode(cu05,null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)) cuname" & _
               ",cp09,cp10,tm44,tm45" & _
               " from trademark,customer,CaseProgress" & _
               " where tm01='" & pa(1) & "' and tm02='" & pa(2) & "' and tm03='" & pa(3) & "' and tm04='" & pa(4) & "'" & _
               " and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+)" & _
               " and cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
               " and cp10='101'"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strTM05 = "" & adoRst.Fields("tm05")
      strTM09 = "" & adoRst.Fields("tm09")
      strApp1Addr = "" & adoRst.Fields("App1Addr")
      strCuName = "" & adoRst.Fields("cuname")
      strCP09 = "" & adoRst.Fields("CP09")
      strCP10 = "" & adoRst.Fields("CP10")
      'Add By Sindy 2025/2/21
      strTM44 = "" & adoRst.Fields("tm44")
      strTM45 = "" & adoRst.Fields("tm45")
      '2025/2/21 END
   End If
   '商標處經理
   strExc(0) = "select st01 from staff" & _
               " where st03='P20' and st20='41' and st04='1'"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   strP20Mgr = ""
   If intI = 1 Then
      strP20Mgr = "" & adoRst.Fields("st01")
   End If
   
   'Modify By Sindy 2018/5/11
   '*************************************************************************************************
   'E-Mail呼叫 frm880019:要將寄信的內容及寄信的成功時間儲存在資料庫中，便於事後查詢。
   lblSendMailDt.Caption = "寄件日期:"
   lblSendMailDt.Visible = True
   frm880019.m_bolSaveMail = True
   frm880019.lblSender = "tm@taie.com.tw" 'Add By Sindy 2018/5/25 寄信者要掛TM信箱
   frm880019.m_CP01 = pa(1)
   frm880019.m_CP02 = pa(2)
   frm880019.m_CP03 = pa(3)
   frm880019.m_CP04 = pa(4)
   frm880019.m_CP09 = strCP09
   frm880019.m_CP10 = strCP10
   '主旨
   'Modify By Sindy 2019/9/9
   'strCaseNo = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "")
   strCaseNo = pa(1) & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "")
   '2019/9/9 END
   'Add By Sindy 2025/2/21 有代理人及彼所案號 -- 天雲
   strExc(10) = ""
   If strTM44 <> "" And strTM45 <> "" Then
      strExc(10) = "貴方卷號:" & strTM45 & " "
   End If
   '2025/2/21 END
   frm880019.txtSubject = "查名(" & Pub_StrUserSt17 & "-" & strCP14ST17 & ")商標名稱:『" & strTM05 & "』(類別:第" & strTM09 & "類 " & strExc(10) & "本所案號:" & strCaseNo & ")"
   '本文
   strGoods = BeforePrintGetDBData("TMGoods:" & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & "-中文含第類")
   strContent = "敬啟者：" & vbCrLf & _
                "　　請依下列事項代為申請商標網上查詢 (包括申請中商標), 並請將查詢結果及近似之商標圖樣儘速e - mial予本所：" & vbCrLf & _
                "一、申請人：" & strCuName & vbCrLf & _
                "　　地址：" & strApp1Addr & vbCrLf & _
                "二、商標名稱：『" & strTM05 & "』" & vbCrLf & _
                "三、商品類別" & vbCrLf & vbCrLf & _
                strGoods & vbCrLf & vbCrLf & _
                "(" & strCaseNo & ")" & vbCrLf & _
                "請確認商品名稱！" & vbCrLf & vbCrLf & vbCrLf
   'Modify By Sindy 2021/7/28 修改”巨京查名”信件內容
   'strContent = strContent & "煩請於工作日3-5天,儘速將查詢結果e-mail告知,謝謝!" & vbCrLf & vbCrLf & vbCrLf
   strContent = strContent & "客戶急須知道查名結果, 儘速將查詢結果e-mail告知,謝謝!" & vbCrLf & vbCrLf & vbCrLf
   '2021/7/28 END
   'Modify By Sindy 2020/1/13 Mark,改圖片簽名檔
'   strContent = strContent & "台一國際專利商標事務所" & vbCrLf & _
'                             "Tai E International Patent & Law Office" & vbCrLf & _
'                             "104 台灣台北市長安東路2段112號9樓" & vbCrLf & _
'                             "TEL: 886 2 25061023 ext 321" & vbCrLf & _
'                             "FAX: 886 2 25011666" & vbCrLf & _
'                             "Email:tm@taie.com.tw; lawoffice@taie.com.tw" & vbCrLf & _
'                             "URL: https://www.taie.com.tw" & vbCrLf & vbCrLf
'   strContent = strContent & "********************保密警語********************" & vbCrLf & _
'                          "本信件僅授權於指定之收信人取閱之用，信件中可能含有機密性資訊。" & vbCrLf & _
'                          "如果您並非被指定之收信人，任何未經授權而擅自使用此信件所含之機密資訊的行為是被嚴格禁止的。" & vbCrLf & _
'                          "如果您在任何未經授權的情形之下收到本信件，煩請您立即告知原發信人並將此信件回傳至以上地址。" & vbCrLf & _
'                          "謝謝您的合作。"
   frm880019.txtContent = strContent
   '2020/1/13 END
   
   '附件
'   m_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum 'Add By Sindy 2017/1/6 以防止上面寄信時有些檔案會被咬住,後面刪檔會有權限問題
'   KillAttach 'Add By Sindy 2017/3/10
'   bolSelFile = False
'   pFiles = ""
'   For ii = 0 To lstAtt(0).ListCount - 1
'      If lstAtt(0).Selected(ii) Then
'         bolSelFile = True
'         stFileName = lstAtt(0).List(ii)
'         If InStrRev(stFileName, " (") > 0 Then
'            stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
'         End If
'         If InStr(stFileName, "\") = 0 Then
'            If GetAttachFile(stFileName, CInt(m_AttEEP02)) = False Then Exit Sub
'         End If
'         pFiles = pFiles & ";" & stFileName
'      End If
'   Next ii
'   If bolSelFile = False Then
'      Call DownloadAllAttachFile(CInt(m_AttEEP02), 0, pFiles)
'   Else
'      If pFiles <> "" Then pFiles = Mid(pFiles, 2)
'   End If
'   frm880019.SetAttach pFiles
   'frm880019.SetEmail m_PA26, m_PA149, m_PA75
   'Add By Sindy 2015/9/23
'   If m_CP44 = "" Then
'      '抓AB類收文號的代理人，預設最後發文日最大收文號的代理人...同發文作業預設的代理人(AddAgent)
'      '2008/2/21 加聯絡人
'      '2010/2/23 香港案要排除421
'      strExc(0) = "SELECT  CP44,cp116,CP45,Max(nvl(CP27,0)||CP09) Srt FROM CaseProgress" & _
'                  " WHERE CP01='" & m_CP01 & "' AND CP02='" & m_CP02 & "'" & _
'                  " AND CP03='" & m_CP03 & "' AND CP04='" & m_CP04 & "'" & _
'                  " AND CP44 IS NOT NULL AND CP09<'C'" & _
'                  IIf(m_PA09 = "013", " AND CP10<>'421'", "") & _
'                  " Group By CP44,cp116,CP45 Order By Srt desc "
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         m_CP44 = "" & RsTemp.Fields("cp44")
'         m_CP116 = "" & RsTemp.Fields("CP116")
'      End If
'   End If
'   '2015/9/23 END
   '收件者.To
   If m_Y52269FA16Mail = "" Then
      MsgBox "無代表號信箱", vbInformation
   End If
   'Modify By Sindy 2018/5/11 林經理說副本不需要再加入67002.葉雪貞特助
   'Modify By Sindy 2020/1/13 1.副本:江協理(98020) 2.密件副本:林經理
   'Modify By Sindy 2021/11/10 林經理退休,改1.副本：江協理、林承慧；2.密件副本：取消。
   If strSrvDate(1) >= 20211115 Then
      frm880019.SetEmail "", "", m_strCP44, , True, "98020;86048", IIf(strCP14 <> "", ";" & strCP14, "")
   Else
   '2021/11/10 END
      frm880019.SetEmail "", "", m_strCP44, , True, "98020", IIf(strCP14 <> "", ";" & strCP14, "") & IIf(strP20Mgr <> "", ";" & strP20Mgr, "")
   End If
   frm880019.cmdAttach.Visible = True 'False
   frm880019.SetParent Me
   frm880019.Show vbModal
   pbolDone = frm880019.m_bolDone
   Unload frm880019
   '*************************************************************************************************
   'm_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") 'Add By Sindy 2017/1/6 以防止上面寄信時有些檔案會被咬住,後面刪檔會有權限問題
   If pbolDone = True Then  '寄信成功
      '寄件日期
      Dim rsTmp As New ADODB.Recordset
      Dim strTemp As String, strCDate As String, strCTime As String
         
      strSql = "Select *" & _
               " From smailbackup" & _
               " Where smb01='" & strCP09 & "'" & _
               " order by smb02 desc,smb03 desc"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      lblSendMailDt.Visible = False
      If rsTmp.RecordCount > 0 Then
         lblSendMailDt.Visible = True
         strTemp = TAIWANDATE(rsTmp.Fields("smb02"))
         strCDate = Format(strTemp, "###/##/##")
         strTemp = rsTmp.Fields("smb03")
         strCTime = Format(strTemp, "##:##:##")
         lblSendMailDt.Caption = "寄件日期:" & strCDate & " " & strCTime
      End If
      rsTmp.Close
      
      'Add By Sindy 2019/6/27 增加通知已查名的聯絡歷程
      '取得最大序號
      intMaxEEP02 = 0
      strSql = "select eep02 From empelectronprocess where eep01='" & strCP09 & "' order by eep02 desc"
      intI = 1
      CheckOC3
      Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         AdoRecordSet3.MoveFirst
         If AdoRecordSet3.RecordCount > 0 Then
            intMaxEEP02 = AdoRecordSet3.Fields(0)
         End If
      End If
      strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep08) values(" & _
               CNULL(strCP09) & "," & intMaxEEP02 + 1 & ",'" & strUserNum & "'," & _
               CNULL(EMP_查名) & "," & _
               CNULL(strCP14) & "," & _
               strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & ",'已通知大陸代理人查名,詳情請參閱卷宗區')"
      cnnConnection.Execute strSql
      '2019/6/27 END
   End If
   
   Set rsTmp = Nothing
   Set adoRst = Nothing
   Exit Sub
   '2018/5/11 END
   
   
'   '取得郵件範本檔名
'   strTemplatePath = PUB_DownloadOftPath(m_StrUserSt03, "P29", , False)
'
'   '呼叫新郵件：
'   Set objOutLook = CreateObject("Outlook.Application")
'   If Dir(strTemplatePath) <> "" Then
'      Set objMail = objOutLook.CreateItemFromTemplate(strTemplatePath)
'   Else
'      Set objMail = objOutLook.CreateItem(0)
'   End If
'   '副本.cc
'   objMail.cc = "67002" '葉雪貞
'   '收件者.To
'   If m_Y52269FA16Mail = "" Then
'      MsgBox "無代表號信箱", vbInformation
'   Else
'      objMail.To = m_Y52269FA16Mail
'   End If
'   '密件副本.BCC
'   objMail.BCC = strCP14 & ";" & strP20Mgr
'   '主旨.Subject
'   strCaseNo = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "")
'   objMail.Subject = "查名(" & Pub_StrUserSt17 & "-" & strCP14ST17 & ")商標名稱:『" & strTM05 & "』( 類別:第" & strTM09 & "類 本所案號:" & strCaseNo & ")"
'
'   '內文.Body
'   strGoods = BeforePrintGetDBData("TMGoods:" & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & "-中文含第類")
'   strContent = "敬啟者：<BR>" & _
'                "　　請依下列事項代為申請商標網上查詢 (包括申請中商標), 並請將查詢結果及近似之商標圖樣儘速e - mial予本所：<BR>" & _
'                "一、申請人：" & strCuName & "<BR>" & _
'                "　　地址：" & strApp1Addr & "<BR>" & _
'                "二、商標名稱：『" & strTM05 & "』<BR>" & _
'                "三、商品類別<BR><BR>" & _
'                strGoods & "<BR><BR>" & _
'                "(" & strCaseNo & ")<BR>" & _
'                "請確認商品名稱！<BR>"
''   strContent = Replace(strContent, "新細明體", "Times New Roman")
''   strContent = Replace(strContent, vbCrLf, "<BR>")
''   strContent = Replace(strContent, "  ", "&nbsp;&nbsp;")
'   'objMail.HTMLBody = "<FONT FACE=""Times New Roman"">" & strContent & "</FONT>"
'   objMail.HTMLBody = "<span style=font-size:13.0pt;font-family:""新細明體"">" & strContent & "</span>" & objMail.HTMLBody
'   'objMail.HTMLBody = strContent & objMail.HTMLBody
'   objMail.Display
'
'   Set objMail = Nothing
'   Set objOutLook = Nothing
'   Set adoRst = Nothing
End Sub

'Add By Sindy 2014/8/13
'Modify By Sindy 2014/10/3
'Private Sub cmdFCMail_Click(Index As Integer)
Public Sub cmdFCMail_Click(Index As Integer)
'2014/10/3 END
Dim strContent As String
Dim strEfile As String
Dim m_bolEmail As Boolean, m_bolPlusPaper As Boolean
Dim mstrBillNo As String
Dim strOutlookText As String, strReport As String, strPath As String 'Add By Sindy 2021/3/29
   
   Call GetStrCustomer 'Added by Lydia 2020/09/11 抓申請人1的編號
   
   'Added by Lydia 2015/10/30
   
   If Index = 2 Then
      'Mark by Lydia 2025/03/13 已不再使用
'        'Added by Lydia 2015/10/30
'        'Modified by Lydia 2016/07/01
'        'Move by Lydia 2022/09/28 從Form_Load搬過來
'        strExc(5) = Pub_GetSpecMan("外專對外翻聯絡人員")
'        'If Pub_StrUserSt03 = "M51" Or InStr("73023,82045", strUserNum) > 0 Then
'        If Pub_StrUserSt03 = "M51" Or InStr(strExc(5), strUserNum) > 0 Then
'        'end 2016/07/01
'           cmdFCMail(2).Visible = True
'           m2FileName = "外專翻譯_案件工作確認單樣本.doc"
'           Call PUB_GetSampleFile(m2FileName, "M51-000299-0-01")
'        End If
'        'end 2022/09/28
'
'       'Added by Lydia 2018/06/25 建立資料夾
'       strSavePath1 = App.path & "\FCmail2"
'       If Dir(strSavePath1, vbDirectory) = "" Then
'            MkDir strSavePath1
'       End If
'       'Added by Lydia 2019/08/06 對外信件要加信尾(郵件範本)
'        If Dir(App.path & "\$$TOT-000F22-0-01.oft") = "" Then
'            Call PUB_GetSampleFile("$$TOT-000F22-0-01.oft", "TOT-000F22-0-01")
'        End If
'
'      'Modified by Lydia 2015/11/11 + CP09
'      'Modified by Lydia 2016/01/11 改抓最新的新案翻譯或其他翻譯927
'      'strSql = "select CP14,ST18,CP48,CP09 from CaseProgress,staff where cp14=st01(+) and cp10='201' and cp01=" & CNULL(Text1) & " and cp02=" & CNULL(Text2) & " and cp03=" & CNULL(IIf(Text3 = "", "00", Text3)) & " and cp04=" & CNULL(IIf(Text4 = "", "00", Text4))
'      'intI = 1
'      'Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      'If RsTemp.RecordCount = 0 Then
'      '   MsgBox "本案無新案翻譯!!", vbCritical
'      'ElseIf InStr("F5588,F5653", UCase("" & RsTemp(0))) > 0 And Not IsNull(RsTemp(0)) Then
'      'Modified by Lydia 2017/09/28 +迅達F5698
'      'Modified by Lydia 2018/01/04 'F5588','F5653','F5698'=> GetAddStr(外翻_舜禹 & "," & 外翻_捷恩凱 & "," & 外翻_迅達)
'      'Modified by Lydia 2018/06/25 判斷是否輸入外文本和圖示頁數
'      'strSql = "select CP14,ST18,CP48,CP09 from CaseProgress,staff where cp14=st01(+) and cp10 in ('201','927') and cp14 in (" & GetAddStr(外翻_舜禹 & "," & 外翻_捷恩凱 & "," & 外翻_迅達) & ") " & _
'               "and cp01=" & CNULL(Text1) & " and cp02=" & CNULL(Text2) & " and cp03=" & CNULL(IIf(Text3 = "", "0", Text3)) & " and cp04=" & CNULL(IIf(Text4 = "", "00", Text4)) & _
'               " order by cp05 desc"
'      strSql = "select CP14,ST18,CP48,CP09,TF24,TF25 from CaseProgress,staff,TransFee where cp14=st01(+) and cp09=tf01(+) and cp10 in ('201','927') and cp14 in (" & GetAddStr(外翻_舜禹 & "," & 外翻_捷恩凱 & "," & 外翻_迅達) & ") " & _
'               "and cp01=" & CNULL(Text1) & " and cp02=" & CNULL(Text2) & " and cp03=" & CNULL(IIf(Text3 = "", "0", Text3)) & " and cp04=" & CNULL(IIf(Text4 = "", "00", Text4)) & _
'               " order by cp05 desc"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If RsTemp.RecordCount = 0 Then
'          MsgBox "本案無外翻承辦的新案翻譯或其他翻譯!!", vbCritical
'      Else
'          If Val("" & RsTemp.Fields("tf24")) = 0 Then  'Added by Lydia 2018/06/25 判斷未輸入外文本頁數,人工填入
'                strLoadPath = UCase("" & RsTemp(0))
'                Set frm880004.mPreForm = Me
'                frm880004.iStiu = 3
'                frm880004.m_TempList = Text1 & Text2 & IIf(Text3 = "", "0", Text3) & IIf(Text4 = "", "00", Text4) 'Added by Lydia 2018/04/30 P案請回復之前做法,由給舜禹翻譯資料夾帶檔案
'                frm880004.Show vbModal
'          'Added by Lydia 2018/06/25 代入外文本頁數
'          Else
'                mPcnt1 = Val("" & RsTemp.Fields("tf24"))
'                mPcnt2 = Val("" & RsTemp.Fields("tf25"))
'          End If
'          'end 2018/06/25
'
'         'Modified by Lydia 2015/11/11 + CP09
'         Call Translate_SendMail(RsTemp(0), "" & RsTemp(1), "" & RsTemp(2), RsTemp(3))
'      'Else
'      '   MsgBox "本案的新案翻譯承辦人非外翻!!", vbCritical
'      'end 2016/01/11
'      End If
      'end 2025/03/13
      Exit Sub
      
   'Add By Sindy 2016/8/1 檢查專利代理人/申請人是否有上傳平台帳號
   ElseIf Index = 0 Then
      'Add By Sindy 2017/3/7
      If cmdFCMail(0).Caption = "巨京查名郵件" Then
         Call TradeMarkP29Mail
         Exit Sub
      End If
      '2017/3/7 END
      
      strSql = PUB_ChkCustWebExist(pa(1), pa(2), pa(3), pa(4))
      If strSql <> "" Then
         MsgBox strSql & "有上傳平台帳號，若是請款程序請一併上傳帳單！", vbInformation
      End If

   '2016/8/1 END
   End If
   'end 2015/10/30
   
   'Added by Lydia 2020/09/11 增加「各項指示」勾選項，勾選後在執行FC郵件一併另外開啟撰寫信函Word檔；原本FC郵件內的備註和各項指示都不再帶出。
   If Index = 0 Or Index = 1 Then
      If ChkINST.Value = 1 Then
          txtFaxFace.Tag = txtFaxFace.Text
          txtFaxFace = "N"
          Call Command1_Click
          txtFaxFace = txtFaxFace.Tag
      End If
   End If
   'end 2020/09/11
   'Added by Lydia 2023/10/04 FMP案待客戶最終指示相關控管
   If Index = 0 And Pub_StrUserSt03 = "F21" And PUB_ChkFMP970mail("3", pa(1), pa(2), pa(3), pa(4), strSql) = True Then
      If strSql <> "" Then
         MsgBox strSql, vbInformation
      End If
   End If
   'end 2023/10/04
   'Add By Sindy 2015/1/8
   m_MySt(1) = pa(1)
   m_MySt(2) = pa(2)
   m_MySt(3) = pa(3)
   m_MySt(4) = pa(4)
   m_SysKind = CheckSys(pa(1))
   SetLetterSt
   '2015/1/8 END
   
   '內文
   strContent = ""
   'Added by Morgan 2023/4/27
   '申請人當中有X45149或 X4514901 < NIKON CORPORATION尼康股份有限公司>
   '在進行 發FC郵件(工程師署名) 時,建立Email在第一行帶入
   '【CONFIDENTIAL ATTORNEY-CLIENT PRIVILEGED COMMUNICATION】
   If Index = 0 And Me.ActiveControl = cmdFCMail(0) And cmdFCMail(0).Caption = "發FC郵件(工程師署名)" Then
      If InStr(m_CustNo(1) & m_CustNo(2) & m_CustNo(3) & m_CustNo(3) & m_CustNo(5), "X4514900") > 0 Or InStr(m_CustNo(1) & m_CustNo(2) & m_CustNo(3) & m_CustNo(3) & m_CustNo(5), "X4514901") > 0 Then
         strContent = strContent & "【CONFIDENTIAL ATTORNEY-CLIENT PRIVILEGED COMMUNICATION】" & vbCrLf & vbCrLf
      End If
   End If
   'end 2023/4/27
   strContent = strContent & ChgEngDate(strSrvDate(1)) & vbCrLf & vbCrLf
   strContent = strContent & GetContentEnglish("發信對象") & vbCrLf
   strContent = strContent & GetContentEnglish("ATTN") & vbCrLf
   strContent = strContent & GetContentEnglish("RE") & vbCrLf
   strContent = strContent & GetContentEnglish("稱謂") & vbCrLf
   If m_StrUserST03 = "F22" Then
      'Modify By Sindy 2015/5/29 依定稿別作區分
'      If InStr("08,13", m_ET01) > 0 And m_ET01 <> "" Then '不加s
'         strContent = strContent & "Please refer to the attached file and confirm safe receipt by return e-mail." & vbCrLf & vbCrLf
'      Else
'      '2015/5/29 END
'         strContent = strContent & "Please refer to the attached files and confirm safe receipt by return e-mail." & vbCrLf & vbCrLf
'      End If
      'Modify By Sindy 2017/6/28 寄附件用語
      'Modify By Sindy 2018/9/6 修改內文
'      strContent = strContent & "We are very pleased to attach our reporting letter for your reference." & vbCrLf & vbCrLf & _
'                                "Please confirm safe receipt by return email." & vbCrLf & vbCrLf
      strContent = strContent & "Please find attached our reporting letter for your reference." & vbCrLf & vbCrLf & _
                                "Please confirm safe receipt by return email." & vbCrLf & vbCrLf
      '2018/9/6 END
      '2017/6/28 END
      
      'Add By Sindy 2021/3/29 通知函承辦單備註設定
      'Modify By Sindy 2021/8/18
      If pa(1) = "FG" Then
         Call PUB_GetFcpEMPBillSpec(pa(1) & pa(2) & pa(3) & pa(4), "03", _
                      pa(26), pa(8) & "," & pa(58) & "," & pa(59) & "," & pa(65) & "," & pa(66), _
                      , , , , , , , , strReport, "FEB24", strOutlookText)
      Else
      '2021/8/18 END
         Call PUB_GetFcpEMPBillSpec(pa(1) & pa(2) & pa(3) & pa(4), "03", _
                      pa(75), pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30), _
                      , , , , , , , , strReport, "FEB24", strOutlookText)
      End If
      '2021/3/29 END
      
      'Add By Sindy 2015/5/29
      If m_ET01 = "07" Then '07.證書函
'         'Add By Sindy 2020/12/17
'         If frm060317_1_SpecNO(pa(26), pa(75)) = "S1" Then
'            strContent = strContent & "According to Dow's Representation Guide, we are only sending the scanned copy of the Patent Certificate to you." & vbCrLf & vbCrLf
'         'Add By Sindy 2021/3/9
'         ElseIf frm060317_1_SpecNO(pa(26), pa(75)) = "S2" Then
'            strContent = strContent & "The original Patent Certificate will be sent to Ms. Michelle Sympson at ASM America, Inc. via courier in the near future." & vbCrLf & vbCrLf
'         'Add By Sindy 2021/3/29
         If strOutlookText <> "" Then
            strContent = strContent & strOutlookText
         '2021/3/29 END
         Else
         '2020/12/17 END
            'Add By Sindy 2017/10/6 亭妙:份數大於1份時,才要加此段文字
            If Val(m_ET99) > 1 Then
            '2017/10/6 END
               strContent = strContent & "Our further report and the original Patent Certificate will be sent to you by registered mail." & vbCrLf & vbCrLf
            End If
         End If
         '定稿,譯文,年費表
         'Modify By Sindy 2016/7/14
         If Pub_StrUserSt03 = "M51" Then
            strEfile = PUB_Getdesktop & "\" & m_MySt(1) & m_MySt(2) & "Letter(Patent Certificate).pdf"
         Else
         '2016/7/14 END
            'Modified by Lydia 2024/07/22 改用變數
            'strEfile = "\\typing2\fcp_workflow\patent certificate\" & m_MySt(1) & m_MySt(2) & "Letter(Patent Certificate).pdf"
            strEfile = "\\" & strTyping2Path & "\fcp_workflow\patent certificate\" & m_MySt(1) & m_MySt(2) & "Letter(Patent Certificate).pdf"
         End If
         If Dir(strEfile) <> "" Then
            strAttach = strAttach & ";" & strEfile
         End If
         '證書
         'Modified by Lydia 2024/07/22 改用變數
         'strEfile = "\\typing2\fcp_workflow\patent certificate\" & m_MySt(1) & m_MySt(2) & "Patent Certificate.pdf"
         strEfile = "\\" & strTyping2Path & "\fcp_workflow\patent certificate\" & m_MySt(1) & m_MySt(2) & "Patent Certificate.pdf"
         If Dir(strEfile) <> "" Then
            strAttach = strAttach & ";" & strEfile
         End If
         '延長專利的Note
         If pa(160) = "A01N" Or pa(160) = "A61K" Then
            'Modified by Lydia 2024/07/22 改用變數
            'strEfile = "\\typing2\fcp_workflow\patent certificate\Note(Concerning Patent Term Extension).pdf"
            strEfile = "\\" & strTyping2Path & "\fcp_workflow\patent certificate\Note(Concerning Patent Term Extension).pdf"
            If Dir(strEfile) <> "" Then
               strAttach = strAttach & ";" & strEfile
            End If
         End If
         '檢查是否純e化
         strExc(0) = "select GetEmailFlag(PA01||PA02||PA03||PA04) eMail from patent where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            '純e化
            'Modify By Sindy 2021/3/29 + Or strReport = "Y"
            'Modify By Sindy 2022/12/9 + 芳如:除純E化，Ｅ＋郵寄亦請自動帶入二個檔案（專利公報通知信及專利公報）
            If "" & RsTemp.Fields(0) = "e" Or "" & RsTemp.Fields(0) = "E" Or strReport = "Y" Then
               strExc(0) = "select 1 from CaseProgress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                           " and cp10='926' and cp57 is null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               '沒有核對已准專利並且為純e化時才須夾帶”公報及信函定稿”
               'Modify By Sindy 2021/3/29 + Or strReport = "Y"
               If intI = 0 Or strReport = "Y" Then
                  'Modify By Sindy 2021/3/29
                  '抓公告公報卷宗區的客戶函和官方來函
                  strExc(0) = "select casepaperpdf.*,cp10 from CaseProgress,casepaperpdf where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                              " and cp10='1228' and cp57 is null" & _
                              " and cp09=cpp01 and (instr(upper(cpp02),'.CUS.')>0 or instr(upper(cpp02),'.'||cp10||'.PDF')>0) "
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     RsTemp.MoveFirst
                     'Add By Sindy 2021/8/3
                     strPath = App.path & Pub_GetSpecMan("EmpFlowAttPath")
                     If Dir(strPath, vbDirectory) = "" Then
                        MkDir strPath
                     End If
                     '2021/8/3 END
                     strPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum
                     'Add By Sindy 2021/8/3
                     If Dir(strPath, vbDirectory) = "" Then
                        MkDir strPath
                     End If
                     '2021/8/3 END
                     Do While Not RsTemp.EOF
                        'CUS=Letter(Patent Gazette).pdf
                        If InStr(UCase(RsTemp.Fields("cpp02")), ".CUS.") > 0 Then
                           strEfile = strPath & "\" & m_MySt(1) & m_MySt(2) & "Letter(Patent Gazette).pdf"
                        '官方來函=Patent Gazette.pdf
                        ElseIf InStr(UCase(RsTemp.Fields("cpp02")), "." & RsTemp.Fields("cp10") & ".PDF") > 0 Then
                           strEfile = strPath & "\" & m_MySt(1) & m_MySt(2) & "Patent Gazette.pdf"
                        End If
                        If Dir(strEfile) = "" Then
                           'Modified by Morgan 2025/3/28 +CPP19
                           If PUB_GetFtpFile("" & RsTemp.Fields("CPP14"), strEfile, , , , "" & RsTemp.Fields("CPP19") <> "") = False Then
                              Exit Sub
                           End If
                        End If
                        If Dir(strEfile) <> "" Then strAttach = strAttach & ";" & strEfile
                        RsTemp.MoveNext
                     Loop
                  Else
                  '2021/3/29 END
                     'Modified by Lydia 2024/07/22 改用變數
                     'strEfile = "\\typing2\FCP_workflow\FCP\" & Left(m_MySt(2), 3) & "\" & m_MySt(1) & m_MySt(2) & "\公報\" & m_MySt(1) & m_MySt(2) & "Letter(Patent Gazette).pdf"
                     strEfile = "\\" & strTyping2Path & "\FCP_workflow\FCP\" & Left(m_MySt(2), 3) & "\" & m_MySt(1) & m_MySt(2) & "\公報\" & m_MySt(1) & m_MySt(2) & "Letter(Patent Gazette).pdf"
                     If Dir(strEfile) <> "" Then
                        strAttach = strAttach & ";" & strEfile
                     End If
                     'Modified by Lydia 2024/07/22 改用變數
                     'strEfile = "\\typing2\FCP_workflow\FCP\" & Left(m_MySt(2), 3) & "\" & m_MySt(1) & m_MySt(2) & "\公報\" & m_MySt(1) & m_MySt(2) & "Patent Gazette.pdf"
                     strEfile = "\\" & strTyping2Path & "\FCP_workflow\FCP\" & Left(m_MySt(2), 3) & "\" & m_MySt(1) & m_MySt(2) & "\公報\" & m_MySt(1) & m_MySt(2) & "Patent Gazette.pdf"
                     If Dir(strEfile) <> "" Then
                        strAttach = strAttach & ";" & strEfile
                     End If
                  End If
               End If
            End If
         End If
      'Add By Sindy 2015/6/16 公開通知函要夾帶公開公報PDF檔案
      ElseIf m_ET01 = "14" Then '14.公開通知函
         m_strFilePath = PUB_GetEFilePath(m_MySt(1)) & "\" & m_MySt(1) & "\" & Left(m_MySt(2), 3) & "\" & m_MySt(1) & m_MySt(2) & "\"
         strEfile = m_strFilePath & m_MySt(1) & m_MySt(2)
         If m_MySt(3) & m_MySt(4) <> "000" Then
            strEfile = strEfile & m_MySt(3) & m_MySt(4)
         End If
         strEfile = strEfile & EfileNameFCP_14
         If Dir(strEfile) <> "" Then
            strAttach = strAttach & ";" & strEfile
         End If
      '2015/6/16 END
      ElseIf m_ET01 = "04" Then '04.核准函
         '新申請案的核准才要夾帶此PDF
         'If InStr(NewCasePtyList, OutCallCP10) > 0 Then
            'Modify By Sindy 2017/6/14
            strEfile = PUB_GetEFilePath(m_MySt(1)) & "\Notice of Allowance with translation\" & m_MySt(1) & m_MySt(2) & EfileNameFCP_04
            '2017/6/14 END
            If Dir(strEfile) <> "" Then
               strAttach = strAttach & ";" & strEfile
            End If
         'End If
      'Add By Sindy 2015/9/8
      ElseIf m_ET01 = "05" Then '05.公告通知函
         'Modified by Lydia 2024/07/22 改用變數
         'strEfile = "\\typing2\FCP_workflow\FCP\" & Left(m_MySt(2), 3) & "\" & m_MySt(1) & m_MySt(2) & "\公報\" & m_MySt(1) & m_MySt(2) & "Letter(Patent Gazette).pdf"
         strEfile = "\\" & strTyping2Path & "\FCP_workflow\FCP\" & Left(m_MySt(2), 3) & "\" & m_MySt(1) & m_MySt(2) & "\公報\" & m_MySt(1) & m_MySt(2) & "Letter(Patent Gazette).pdf"
         If Dir(strEfile) <> "" Then
            strAttach = strAttach & ";" & strEfile
         End If
         'Modified by Lydia 2024/07/22 改用變數
         'strEfile = "\\typing2\FCP_workflow\FCP\" & Left(m_MySt(2), 3) & "\" & m_MySt(1) & m_MySt(2) & "\公報\" & m_MySt(1) & m_MySt(2) & "Patent Gazette.pdf"
         strEfile = "\\" & strTyping2Path & "\FCP_workflow\FCP\" & Left(m_MySt(2), 3) & "\" & m_MySt(1) & m_MySt(2) & "\公報\" & m_MySt(1) & m_MySt(2) & "Patent Gazette.pdf"
         If Dir(strEfile) <> "" Then
            strAttach = strAttach & ";" & strEfile
         End If
      '2015/9/8 END
      'Add By Sindy 2015/7/9
      ElseIf m_ET01 = "08" Then '08.催繳年費通知函
         m_bolEmail = PUB_GetEMailFlag(m_MySt(1) & m_MySt(2) & m_MySt(3) & m_MySt(4), True, , m_bolPlusPaper)
         '檢查是否有請款單
         Call PUB_GetUnPaidBill(m_MySt(1), m_MySt(2), m_MySt(3), m_MySt(4), mstrBillNo)
         If mstrBillNo <> "" Then
            EfileNameFCP_08 = ""
            '產生E化請款單
            PUB_PrintBill mstrBillNo, "", m_bolEmail, m_bolPlusPaper, Me.Name, , 1, "2"
            If EfileNameFCP_08 <> "" Then
               EfileNameFCP_08 = Mid(EfileNameFCP_08, 2)
               strAttach = strAttach & ";" & EfileNameFCP_08
               '檢查是否有部分收款
               strExc(0) = "select a1k01 from acc1k0 where a1k01='" & mstrBillNo & "' and nvl(a1k30,0)>0 and a1k29 is null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If RsTemp.RecordCount > 0 Then
                     MsgBox "請款單編號（" & mstrBillNo & "）有部分收款！", vbInformation
                  End If
               End If
            End If
         End If
      '2015/7/9 END
      End If
      '2015/5/29 END
      If Index = 0 Then
         '取得郵件範本檔名
         strTemplatePath = PUB_DownloadOftPath(m_StrUserST03, Pub_StrUserSt17, EMailType, False)
      Else
         '取得郵件範本檔名
         strTemplatePath = PUB_DownloadOftPath("F23", "", EMailType, False, IIf(Option1(2).Value = True, "3", "2"))
      End If
   End If
   If m_StrUserST03 = "F23" And Index = 0 Then
      '取得郵件範本檔名
      strTemplatePath = PUB_DownloadOftPath(m_StrUserST03, Pub_StrUserSt17, EMailType, , IIf(Option1(2).Value = True, "3", "2"))
   End If
   If Left(strAttach, 1) = ";" Then strAttach = Mid(strAttach, 2)
   'Modify By Sindy 2017/8/14 + , IIf(Option2.Value = True, "FC", IIf(Option6.Value = True, "CF", ""))
   'Modify By Sindy 2019/5/23 + , IIf(bolFrom1105Callme = True, OutCallProcCP10, "")
   Call PUB_SettingFCeMail(m_StrUserST03, strTemplatePath, EMailType, _
                           pa(1), pa(2), pa(3), pa(4), _
                           strContent, strAttach, IIf(bolFrom1105Callme = True, OutCallCP10, ""), m_ET01, m_ET03, IIf(Option2.Value = True, "FC", IIf(Option6.Value = True, "CF", "")), IIf(bolFrom1105Callme = True, OutCallProcCP10, ""))
End Sub

'Add By Sindy 2014/8/14
'截取信件內容
Private Function GetContentEnglish(strGetItem As String, Optional bolCallMail As Boolean = True, Optional ByRef stCaseNo As String, _
                                   Optional ByRef stYourRef As String, Optional ByRef stOurRef As String) As String
   Dim stLetter As String
   Dim m As Integer, j As Integer
   Dim strCustName As String
   Dim strReceiver As String
   Dim bAllApp As Boolean
   Dim bBASF As Boolean 'Added by Morgan 2018/6/25
   
   GetContentEnglish = ""
   '代理人名成兩欄並一行印
   If fa(1) <> "" And Len(fa(1)) <= 30 Then
      If fa(2) <> "" Then
         fa(1) = Trim(fa(1)) & " " & Trim(fa(2))
         fa(2) = ""
      End If
      If fa(3) <> "" Then
         fa(2) = Trim(fa(3)) & " " & Trim(fa(4))
         fa(3) = ""
         fa(4) = ""
      End If
   End If
   If cfa(1) <> "" And Len(cfa(1)) <= 30 Then
      If cfa(2) <> "" Then
         cfa(1) = Trim(cfa(1)) & " " & Trim(cfa(2))
         cfa(2) = ""
      End If
      If cfa(3) <> "" Then
         cfa(2) = Trim(cfa(3)) & " " & Trim(cfa(4))
         cfa(3) = ""
         cfa(4) = ""
      End If
   End If
   If cu(1) <> "" And Len(cu(1)) <= 30 Then
      If cu(2) <> "" Then
         cu(1) = Trim(cu(1)) & " " & Trim(cu(2))
         cu(2) = ""
      End If
      If cu(3) <> "" Then
         cu(2) = Trim(cu(3)) & " " & Trim(cu(4))
         cu(3) = ""
         cu(4) = ""
      End If
   End If
   
   Select Case UCase(strGetItem)
      Case "發信對象"
         'Add By Sindy 2015/1/8
         If bolFrom1105Callme = True Then
            GetContentEnglish = ExceptFieldData("信函抬頭/英", OutCallCP10) & vbCrLf
            
            '外專人員要帶出案件備註及代理人備註
'            'Remove by Lydia 2020/09/11 原本FC郵件內的備註和各項指示都不再帶出
'            If Left(m_StrUserST03, 2) = "F2" Then
'               If Text5 <> "" Then
'                  GetContentEnglish = GetContentEnglish & vbCrLf
'                  'Modified by Lydia 2020/06/04 備註和各項指示並存(暫時到完全取代)
'                  'GetContentEnglish = GetContentEnglish & "案件備註：" & Text5 & vbCrLf
'                  'Modified by Lydia 2020/07/27 備註及各項指示的標題單獨一行=> "：" & vbCrLf
'                  GetContentEnglish = GetContentEnglish & "案件備註：" & vbCrLf & IIf(strKeyCase <> "", Replace(Text5, vbCrLf & strMemoCase, ""), Text5) & vbCrLf
'                  If strKeyCase <> "" Then
'                      GetContentEnglish = GetContentEnglish & "案件各項指示：" & vbCrLf & strMemoCase & vbCrLf
'                  End If
'                  'end 2020/06/04
'               End If
'               If Text6(0) <> "" Then
'                  GetContentEnglish = GetContentEnglish & vbCrLf
'                  'Modified by Lydia 2020/06/04 備註和各項指示並存(暫時到完全取代)
'                  'GetContentEnglish = GetContentEnglish & "代理人備註：" & Text6(0) & vbCrLf
'                  'Modified by Lydia 2020/07/27 備註及各項指示的標題單獨一行=> "：" & vbCrLf
'                  GetContentEnglish = GetContentEnglish & "代理人備註：" & vbCrLf & IIf(strKeyY <> "", Replace(Text6(0), vbCrLf & strMemoY, ""), Text6(0)) & vbCrLf
'                  If strKeyY <> "" Then
'                      GetContentEnglish = GetContentEnglish & "代理人各項指示：" & vbCrLf & strMemoY & vbCrLf
'                  End If
'                  'end 2020/06/04
'               End If
'               If Text6(1) <> "" Then
'                  GetContentEnglish = GetContentEnglish & vbCrLf
'                  'Modified by Lydia 2020/06/04 備註和各項指示並存(暫時到完全取代)
'                  'GetContentEnglish = GetContentEnglish & "申請人備註：" & Text6(1) & vbCrLf
'                  'Modified by Lydia 2020/07/27 備註及各項指示的標題單獨一行=> "：" & vbCrLf
'                  GetContentEnglish = GetContentEnglish & "申請人備註：" & vbCrLf & IIf(strKeyX <> "", Replace(Text6(1), vbCrLf & strMemoX, ""), Text6(1)) & vbCrLf
'                  If strKeyX <> "" Then
'                      GetContentEnglish = GetContentEnglish & "申請人各項指示：" & vbCrLf & strMemoX & vbCrLf
'                  End If
'                  'end 2020/06/04
'               End If
'            End If
'            'end 2020/09/11
         Else
         '2015/1/8 END
            '發信對象為FC代理人
            If Option2.Value Then
               For m = 1 To 4
                  If fa(m) <> "" Then
                     GetContentEnglish = GetContentEnglish & fa(m) & vbCrLf
                     '帶出英文代理人名稱
                     j = m
                     GoTo a
                  ElseIf fa(m) = "" Then
                     GetContentEnglish = GetContentEnglish & "                                         " & vbCrLf
                  End If
                  GoTo B
               Next
a:
               While j <= 3
                  If fa(j + 1) <> "" Then
                     GetContentEnglish = GetContentEnglish & fa(j + 1) & vbCrLf
                  End If
                  j = j + 1
               Wend
B:
               For m = 17 To 22
                  If fa(m) <> "" Then
                     GetContentEnglish = GetContentEnglish & fa(m) & vbCrLf
                  End If
               Next
               
               '外專人員要帶出案件備註及代理人備註
               '調整順序改先案件備註後代理人備註(原來相反)--黃美珍
'               'Remove by Lydia 2020/09/11 原本FC郵件內的備註和各項指示都不再帶出
'               If bolCallMail = True Then
'                  If Left(m_StrUserST03, 2) = "F2" Then
'                     If Text5 <> "" Then
'                        GetContentEnglish = GetContentEnglish & vbCrLf
'                        'Modified by Lydia 2020/06/04 備註和各項指示並存(暫時到完全取代)
'                        'GetContentEnglish = GetContentEnglish & "案件備註：" & Text5 & vbCrLf
'                        'Modified by Lydia 2020/07/27 備註及各項指示的標題單獨一行=> "：" & vbCrLf
'                        GetContentEnglish = GetContentEnglish & "案件備註：" & vbCrLf & IIf(strKeyCase <> "", Replace(Text5, vbCrLf & strMemoCase, ""), Text5) & vbCrLf
'                        If strKeyCase <> "" Then
'                            GetContentEnglish = GetContentEnglish & "案件各項指示：" & vbCrLf & strMemoCase & vbCrLf
'                        End If
'                        'end 2020/06/04
'                     End If
'                     If Text6(0) <> "" Then
'                        GetContentEnglish = GetContentEnglish & vbCrLf
'                        'Modified by Lydia 2020/06/04 備註和各項指示並存(暫時到完全取代)
'                        'GetContentEnglish = GetContentEnglish & "代理人備註：" & Text6(0) & vbCrLf
'                        'Modified by Lydia 2020/07/27 備註及各項指示的標題單獨一行=> "：" & vbCrLf
'                        GetContentEnglish = GetContentEnglish & "代理人備註：" & vbCrLf & IIf(strKeyY <> "", Replace(Text6(0), vbCrLf & strMemoY, ""), Text6(0)) & vbCrLf
'                        If strKeyY <> "" Then
'                            GetContentEnglish = GetContentEnglish & "代理人各項指示：" & vbCrLf & strMemoY & vbCrLf
'                        End If
'                        'end 2020/06/04
'                     End If
'                     If Text6(1) <> "" Then
'                        GetContentEnglish = GetContentEnglish & vbCrLf
'                        'Modified by Lydia 2020/06/04 備註和各項指示並存(暫時到完全取代)
'                        'GetContentEnglish = GetContentEnglish & "代理人備註：" & Text6(1) & vbCrLf
'                        'Modified by Lydia 2020/07/27 備註及各項指示的標題單獨一行=> "：" & vbCrLf
'                        GetContentEnglish = GetContentEnglish & "申請人備註：" & vbCrLf & IIf(strKeyX <> "", Replace(Text6(1), vbCrLf & strMemoX, ""), Text6(1)) & vbCrLf
'                        If strKeyX <> "" Then
'                            GetContentEnglish = GetContentEnglish & "申請人各項指示：" & vbCrLf & strMemoX & vbCrLf
'                        End If
'                        'end 2020/06/04
'                     End If
'                  End If
'               End If
'               'end 2020/09/11
               
            '發信對象為CF代理人
            ElseIf Option6.Value Then
               For m = 1 To 4
                  If cfa(m) <> "" Then
                     GetContentEnglish = GetContentEnglish & cfa(m) & vbCrLf
                     '帶出英文代理人名稱
                     j = m
                     GoTo h
                  ElseIf cfa(m) = "" Then
                     GetContentEnglish = GetContentEnglish & "                                         " & vbCrLf
                  End If
                  GoTo i
               Next
h:
               While j <= 4
                  If cfa(j + 1) <> "" Then
                     GetContentEnglish = GetContentEnglish & cfa(j + 1) & vbCrLf
                  End If
                  j = j + 1
               Wend
i:
               For m = 17 To 22
                  If cfa(m) <> "" Then
                     GetContentEnglish = GetContentEnglish & cfa(m) & vbCrLf
                  End If
               Next
               
            '發信對象為申請人
            ElseIf Option3.Value Then
               For m = 1 To 4
                  If cu(m) <> "" Then
                     GetContentEnglish = GetContentEnglish & cu(m) & vbCrLf
                     '帶出英文申請人名稱
                     j = m
                     GoTo c
                  ElseIf cu(m) = "" Then
                     GetContentEnglish = GetContentEnglish & "                                         " & vbCrLf
                  End If
                  GoTo D
               Next
c:
               While j <= 4
                  If cu(j + 1) <> "" Then
                     GetContentEnglish = GetContentEnglish & cu(j + 1) & vbCrLf
                  End If
                  j = j + 1
               Wend
D:
               For m = 17 To 22
                  If cu(m) <> "" Then
                     GetContentEnglish = GetContentEnglish & cu(m) & vbCrLf
                  End If
               Next
            End If
         End If
      Case "ATTN"
         Combo4.Text = Trim(Combo4.Text)
         Combo5.Text = Trim(Combo5.Text)
         'Add By Sindy 2015/9/14 08.繳年費通知函的聯絡人要抓取年費連絡人
         'Add By Sindy 2020/12/8 ex:FCP-58532 實體審查-08.繳年費通知函
         'If bolFrom1105Callme = True And m_ET01 = "08" Then
         If bolFrom1105Callme = True And (OutCallCP10 = "605" Or OutCallCP10 = "1605") Then
         '2020/12/8 END
            GetContentEnglish = ExceptFieldData("聯絡人1/英", OutCallCP10) & vbCrLf
         Else
         '2015/9/14 END
            '發信對象為FC代理人
            If Option2.Value Then
               'Added by Morgan 2024/5/14
               '外專程序發FC郵件且定稿語文為日文時，聯絡人先帶英文名，沒有再帶日文名 -- Arashi
               If bolCallMail And m_StrUserST03 = "F22" And m_iLang = "2" And m_strContact1(1) <> "" Then
                  If m_strContact2(1) <> "" Then
                     GetContentEnglish = GetContentEnglish & "Attn: " & m_strContact1(1) & " and " & m_strContact2(1) & vbCrLf
                  Else
                     GetContentEnglish = GetContentEnglish & "Attn: " & m_strContact1(1) & vbCrLf
                  End If
               Else
               'end 2024/5/14
               
                  If Option4.Value And (Combo4 <> "" Or Combo5 <> "") Then
                     If Combo4 <> "" And Combo5 <> "" Then
                        GetContentEnglish = GetContentEnglish & "Attn: " & Combo4 & " and " & Combo5 & vbCrLf
                     Else
                        GetContentEnglish = GetContentEnglish & "Attn: " & Combo4 & Combo5 & vbCrLf
                     End If
                  ElseIf Option5.Value And Combo6 <> "" Then
                     GetContentEnglish = GetContentEnglish & "Attn: " & Combo6 & vbCrLf
                  End If
                  
               End If
               
            '發信對象為CF代理人
            ElseIf Option6.Value Then
               If Option4.Value And (Combo4 <> "" Or Combo5 <> "") Then
                  If Combo4 <> "" And Combo5 <> "" Then
                     GetContentEnglish = GetContentEnglish & "Attn: " & Combo4 & " and " & Combo5 & vbCrLf
                  Else
                     GetContentEnglish = GetContentEnglish & "Attn: " & Combo4 & Combo5 & vbCrLf
                  End If
               ElseIf Option5.Value And Combo6 <> "" Then
                  GetContentEnglish = GetContentEnglish & "Attn: " & Combo6 & vbCrLf
               End If
               
            '發信對象為申請人
            ElseIf Option3.Value Then
               If Option4.Value And (Combo4 <> "" Or Combo5 <> "") Then
                  If Combo4 <> "" And Combo5 <> "" Then
                     GetContentEnglish = GetContentEnglish & "Attn: " & Combo4 & " and " & Combo5 & vbCrLf
                  Else
                     GetContentEnglish = GetContentEnglish & "Attn: " & Combo4 & Combo5 & vbCrLf
                  End If
               ElseIf Option5.Value And Combo6 <> "" Then
                  '帶出聯絡人
                  GetContentEnglish = GetContentEnglish & "Attn: " & Combo6 & vbCrLf
               End If
            End If
         End If
         
      Case "RE"
         stCaseNo = "": stYourRef = "": stOurRef = ""
         '若為專利案件
         If pa(1) = "P" Or pa(1) = "CFP" Or pa(1) = "FCP" Then
            GetContentEnglish = "Re:"
            '改抓函數(比照傳真封面)
            If pa(1) = "P" Or pa(1) = "FCP" Then
               stCaseNo = stCaseNo & " " & PUB_GetNationEngNameForLet(pa(9))
            End If
            stCaseNo = stCaseNo & PUB_GetEngPatKindName("1", pa(8))
            'FCP比照傳真封面
            If pa(1) = "FCP" Then
               stCaseNo = stCaseNo & " Patent Application No. " & pa(11)
               If pa(22) <> "" Then
                  stCaseNo = stCaseNo & " (Patent No. " & pa(22) & ")"
               End If
            Else
               '若有專利號數
               If pa(22) <> "" Then
                  stCaseNo = stCaseNo & " Patent No. " & pa(22)
               '若無專利號數有申請案號
               ElseIf pa(22) = "" And pa(11) <> "" Then
                  stCaseNo = stCaseNo & " Patent Application No. " & pa(11)
               '若無專利號數也無申請案號
               ElseIf pa(22) = "" And pa(11) = "" Then
                  stCaseNo = stCaseNo & " Patent Application No."
               End If
            End If
            GetContentEnglish = GetContentEnglish & stCaseNo & vbCrLf
            
            '申請人名稱
            '日本印中文
            If pa(1) = "FCP" Then
               bAllApp = True
            Else
               '帶出所有申請人
               bAllApp = True
            End If
            If Label11 = "011" Then
               strCustName = GetCustName(Text1.Text & Text2.Text & Text3.Text & Text4.Text, "1", bAllApp, String(11, " "), True)
            Else
               strCustName = GetCustName(Text1.Text & Text2.Text & Text3.Text & Text4.Text, "2", bAllApp, String(11, " "), True)
            End If
            If strCustName <> "" Then
               GetContentEnglish = GetContentEnglish & IIf(bolCallMail = True, "   ", "") & "   Applicant: " & strCustName & vbCrLf
            End If
            '案件名稱
            'Add By Sindy 2015/5/29
            'If bolFrom1105Callme = False Then '定稿維護時不顯示Title
            '2015/5/29 END
               'Modify By Sindy 2018/11/15 外專承辦人洪培堯反應要顯示Title,Matter ID
               If pa(6) <> "" Then
                  GetContentEnglish = GetContentEnglish & IIf(bolCallMail = True, "   ", "") & "   Title: " & SplitTitle(pa(6), IIf(bolCallMail = True, 14, 8)) & vbCrLf
               End If
               If pa(159) <> "" And Left(pa(75), 6) = "Y54047" Then
                  GetContentEnglish = GetContentEnglish & IIf(bolCallMail = True, "   ", "") & "   Matter No.: " & pa(159) & vbCrLf
               End If
               '2018/11/15 END
            'End If
            If Option6.Value = False Then
               '客戶案件案號
               If pa(48) <> "" Then
                  GetContentEnglish = GetContentEnglish & IIf(bolCallMail = True, "   ", "") & "   Case No: " & pa(48) & vbCrLf
               End If
            End If
            
            '彼所案號
            'FC
            If Option2.Value = True Then
               'Modify By Sindy 2019/5/23
               'stYourRef = IIf(bolCallMail = True, "   ", "") & "   Your Ref: " & pa(77)
               stYourRef = IIf(bolCallMail = True, "   ", "") & "   Your Ref: " & IIf(OutCallCP10 = "605", IIf(pa(106) <> "" And pa(76) <> "", pa(106), pa(77)), pa(77))
               '2019/5/23 END
               GetContentEnglish = GetContentEnglish & stYourRef & vbCrLf
            'CF
            ElseIf Option6.Value = True Then
               'Modified by Morgan 2024/3/27
               'strExc(0) = "select CP45 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
               '                " AND CP44 Is Not Null AND cp57 is null AND CP09<'C' Order By CP27 Desc, CP09 Desc "
               'cp(0) = ""
               'intI = 1
               'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               'If intI = 1 Then
               '   With RsTemp
               '      If Not IsNull(.Fields(0)) Then cp(0) = .Fields(0)
               '   End With
               'End If
               'stYourRef = IIf(bolCallMail = True, "   ", "") & "   Your Ref: " & cp(0)
               stYourRef = IIf(bolCallMail = True, "   ", "") & "   Your Ref: " & m_strCP45
               'end 2024/3/27
               
               GetContentEnglish = GetContentEnglish & stYourRef & vbCrLf
            End If
            
         ElseIf pa(1) = "FG" Then
            GetContentEnglish = "Re:"
            stCaseNo = " " & pa(6)
            GetContentEnglish = GetContentEnglish & stCaseNo & vbCrLf
            '申請人名稱
            '日本印中文
            If Label11 = "011" Then
               strCustName = GetCustName(Text1.Text & Text2.Text & Text3.Text & Text4.Text, "1")
            Else
               strCustName = GetCustName(Text1.Text & Text2.Text & Text3.Text & Text4.Text, "2")
            End If
            If strCustName <> "" Then
               GetContentEnglish = GetContentEnglish & IIf(bolCallMail = True, "   ", "") & "   Applicant: " & strCustName & vbCrLf
            End If
            '客戶案件案號
            If pa(29) <> "" Then
               GetContentEnglish = GetContentEnglish & IIf(bolCallMail = True, "   ", "") & "   Case No: " & pa(29) & vbCrLf
            End If
            '彼所案號
            'FC
            If Option2.Value = True Then
               stYourRef = IIf(bolCallMail = True, "   ", "") & "   Your Ref: " & pa(27)
               GetContentEnglish = GetContentEnglish & stYourRef & vbCrLf
            'CF
            ElseIf Option6.Value = True Then
               strExc(0) = "select CP45 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP44 Is Not Null AND CP57 is null AND CP09<'C' Order By CP27 Desc, CP09 Desc "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strExc(1) = "" & RsTemp.Fields(0)
               End If
               stYourRef = IIf(bolCallMail = True, "   ", "") & "   Your Ref: " & strExc(1)
               GetContentEnglish = GetContentEnglish & stYourRef & vbCrLf
            End If
           
         '若為商標案件
         ElseIf pa(1) = "T" Or pa(1) = "TF" Or pa(1) = "CFT" Or pa(1) = "FCT" Then
            GetContentEnglish = "Re:"
            '有審定號
            If pa(15) <> "" Then
               If pa(1) = "CFT" Then
                  stCaseNo = " Trademark Registration No. " & pa(15) & " in " & GetNationName(pa(10), 1)
               Else
                  stCaseNo = " Taiwan Trademark Registration No. " & pa(15)
               End If
            '無審定號印申請案號
            Else
               If pa(1) = "CFT" Then
                  stCaseNo = " Trademark Application No. " & pa(12) & " in " & GetNationName(pa(10), 1)
               Else
                  stCaseNo = " Taiwan Trademark Application No. " & pa(12)
               End If
            End If
            GetContentEnglish = GetContentEnglish & stCaseNo & vbCrLf
g:
            '申請人名稱
            strCustName = GetCustName(Text1.Text & Text2.Text & Text3.Text & Text4.Text, "2")
            If strCustName <> "" Then
               GetContentEnglish = GetContentEnglish & IIf(bolCallMail = True, "   ", "") & "   Applicant: " & strCustName & vbCrLf
            End If
            '案件名稱
            If pa(5) & pa(6) & pa(7) <> "" Then
               If pa(1) = "CFT" Or pa(1) = "FCT" Then
                  GetContentEnglish = GetContentEnglish & IIf(bolCallMail = True, "   ", "") & "   Mark: " & pa(5) & pa(6) & pa(7) & vbCrLf
               Else
                  'Add By Sindy 2015/5/29
                  If bolFrom1105Callme = False Then '定稿維護時不顯示Title
                  '2015/5/29 END
                     GetContentEnglish = GetContentEnglish & IIf(bolCallMail = True, "   ", "") & "   Title: " & pa(5) & pa(6) & pa(7) & vbCrLf
                  End If
               End If
            End If
            '客戶案件案號
            If pa(35) <> "" Then
               GetContentEnglish = GetContentEnglish & IIf(bolCallMail = True, "   ", "") & "   Case No: " & pa(35) & vbCrLf
            End If
            '彼所案號
            'FC
            If Option2.Value = True Then
               stYourRef = IIf(bolCallMail = True, "   ", "") & "   Your Ref: " & pa(45)
               GetContentEnglish = GetContentEnglish & stYourRef & vbCrLf
            'CF
            ElseIf Option6.Value = True Then
               strExc(0) = "select CP45 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                               " AND CP44 Is Not Null AND cp57 is null AND CP09<'C' Order By CP27 Desc, CP09 Desc "
               cp(0) = ""
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  With RsTemp
                     If Not IsNull(.Fields(0)) Then cp(0) = .Fields(0)
                  End With
               End If
               stYourRef = IIf(bolCallMail = True, "   ", "") & "   Your Ref: " & cp(0)
               GetContentEnglish = GetContentEnglish & stYourRef & vbCrLf
            End If
            
         '若為法務案件
         'modify by sonia 2019/7/30 +ACS系統類別
         ElseIf pa(1) = "FCL" Or pa(1) = "CFL" Or pa(1) = "LIN" Or pa(1) = "ACS" Then
            GetContentEnglish = "Re:" & vbCrLf
            '彼所案號
            'FC
            If Option2.Value = True Then
               stYourRef = IIf(bolCallMail = True, "   ", "") & "   Your Ref: " & pa(23)
               GetContentEnglish = GetContentEnglish & stYourRef & vbCrLf
            'CF
            ElseIf Option6.Value = True Then
               strExc(0) = "select CP45 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                               " AND CP44 Is Not Null AND CP57 is null AND CP09<'C' Order By CP27 Desc, CP09 Desc "
               cp(0) = ""
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  With RsTemp
                     If Not IsNull(.Fields(0)) Then cp(0) = .Fields(0)
                  End With
               End If
               stYourRef = IIf(bolCallMail = True, "   ", "") & "   Your Ref: " & cp(0)
               GetContentEnglish = GetContentEnglish & stYourRef & vbCrLf
            End If
         End If
         '本所案號
         stOurRef = IIf(bolCallMail = True, "   ", "") & "   Our Ref: " & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & IIf(pa(4) = "00", "", "-" & pa(4)))
         GetContentEnglish = GetContentEnglish & stOurRef & vbCrLf
         
         'Added by Morgan 2017/7/28 外專工程師撰寫代理人為BASF(Y45814000)的信函時要帶出 Literature cited,若有1202審查意見通知函或1002核駁未發文時後面帶 Yes/No 否則帶 No
         'Modified by Morgan 2018/6/25 +外專承辦也要(F23)
         'If m_StrUserSt03 = "F21" Then
         bBASF = False
         If m_StrUserST03 = "F21" Or m_StrUserST03 = "F23" Then
         'end 2018/6/25
            m_MySt(1) = pa(1): m_MySt(2) = pa(2): m_MySt(3) = pa(3): m_MySt(4) = pa(4)
            'Added by Morgan 2018/12/7
            m_SysKind = CheckSys(pa(1))
            SetLetterSt
            'end 2018/12/7
            'Modified by Morgan 2018/8/31 +Y54554 BASF (China)
            'Modified by Morgan 2018/11/30 改呼要函數以便新增編號時此處不必再改
            'If (pa(75) = "Y45814" Or pa(75) = "Y54554") Then
            If ExceptFieldData2("BASF要印") <> "" Then
            'end 2018/11/30
               'Added by Morgan 2018/6/25
               bBASF = True
               GetContentEnglish = GetContentEnglish & "   Official Deadline: " & ExceptFieldData2("下一程序法限/英") & vbCrLf
               'Modify By Sindy 2021/7/20
               If Not (pa(1) = "FCP" And Val(Right(Combo8, 4)) = "926") Then 'Added by Morgan 2024/4/1 FCP的二核報告不要帶
                  If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
                     'Modified by Morgan 2021/8/2
                     'GetContentEnglish = GetContentEnglish & "   Tai E Working Deadline: " & ExceptFieldData2("下一程序約定期限/英") & vbCrLf
                     GetContentEnglish = GetContentEnglish & "   Tai E Alert Date: " & ExceptFieldData2("下一程序約定期限/英") & vbCrLf
                     'end 2021/8/2
                  Else
                  '2021/7/20 END
                     GetContentEnglish = GetContentEnglish & "   Tai E Working Deadline: " & ExceptFieldData2("下一程序所限/英") & vbCrLf
                  End If
               End If
               GetContentEnglish = GetContentEnglish & "   Name of First Inventor: " & ExceptFieldData2("發明人1/英") & vbCrLf
               'end 2018/6/25
               
               strExc(0) = "select cp09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND cp10 in ('1202','1002') AND CP57||cp27 is null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  GetContentEnglish = GetContentEnglish & "   Literature cited: Yes/No" & vbCrLf
                  GetContentEnglish = GetContentEnglish & "   New Matter: Yes/No" & vbCrLf 'Added by Morgan 2023/8/9
               Else
                  GetContentEnglish = GetContentEnglish & "   Literature cited: No" & vbCrLf
                  GetContentEnglish = GetContentEnglish & "   New Matter: No" & vbCrLf 'Added by Morgan 2023/8/9
               End If
            End If
            
            'Added by Morgan 2025/4/24
            If ExceptFieldData3("SE要印") <> "" Then
               GetContentEnglish = GetContentEnglish & "   Filing Date: " & TranslateKeyWord(incCNV_ENGLISH_DATE, DBDATE(pa(10)), "英") & vbCrLf
               GetContentEnglish = GetContentEnglish & "   Priority Appl. No: " & ExceptFieldData2("優先權國家+申請號/英") & vbCrLf
               GetContentEnglish = GetContentEnglish & "   Priority Date: " & ExceptFieldData2("優先權日/英") & vbCrLf
               GetContentEnglish = GetContentEnglish & "   Parent Appl. No: " & ExceptFieldData("分割母案申請案號N/A") & vbCrLf
            End If
            'end 2025/4/24
            
            'Added by Morgan 2019/6/28 --潘子微
            '瑞典代理人Y34412 SCANIA CV AB PATENTS 之工程師撰寫信函其信頭列出該案已請款之總金額USD
            If m_StrUserST03 = "F21" And pa(75) = "Y34412" Then
               GetContentEnglish = GetContentEnglish & "   Total amount billed so far for this case (USD): " & GetBillUSAmount(pa(1), pa(2), pa(3), pa(4)) & vbCrLf
            End If
            'end 2019/6/28
         End If
         'end 2017/7/28
         
         'FG也不要--David
         'CFT, FCT也不要--陳經理
         'If pa(1) <> "CFP" Then
         'Add By Sindy 2015/5/29
         'Modify By Sindy 2015/7/16
         'If bolFrom1105Callme = False Then '定稿維護時不顯示Title
         'Modified by Morgan 2018/6/25 外專 BASF 案件統一上面帶特殊資訊
         'If m_StrUserSt03 <> "F22" Then
         'Modified by Morgan 2024/4/1 FCP的二核報告不要帶
         'If m_StrUserST03 <> "F22" And Not bBASF Then
         If m_StrUserST03 <> "F22" And Not bBASF And Not (pa(1) = "FCP" And Val(Right(Combo8, 4)) = "926") Then
         'end 2024/4/1
         'end 2018/6/25
         '2015/7/16 END
         '2015/5/29 END
            'Added by Morgan 2021/2/2 寶齡富錦 Y55435 特殊控制
            If bIsBPFCase Then
               If stNP09 <> "" Then
                  GetContentEnglish = GetContentEnglish & IIf(bolCallMail = True, "   ", "") & "   Deadline: " & ChgEngDate(stNP09) & vbCrLf
               Else
                  GetContentEnglish = GetContentEnglish & IIf(bolCallMail = True, "   ", "") & "   Deadline: " & vbCrLf
               End If
            Else
            'end 2021/2/2
                  
               '增加LIN系統類別
               'modify by sonia 2019/7/30 +ACS系統類別
               If pa(1) <> "CFP" And pa(1) <> "CFT" And pa(1) <> "FCT" And _
                  pa(1) <> "FCL" And pa(1) <> "CFL" And pa(1) <> "FG" And pa(1) <> "LIN" And pa(1) <> "ACS" Then
                  '下一程序期限
                  strExc(0) = GetNPSQL
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     With RsTemp
                        'Modified by Morgan 2021/8/2
                        'If Not IsNull(.Fields(0)) Then np(0) = .Fields(0)
                        If Not IsNull(.Fields("np23")) Then np(0) = .Fields("np23")
                        'end 2021/8/2
                     End With
                     '改印英文日期-- David
                     'Modified by Morgan 2021/8/2
                     'GetContentEnglish = GetContentEnglish & IIf(bolCallMail = True, "   ", "") & "   Deadline: " & ChgEngDate(np(0)) & vbCrLf
                     GetContentEnglish = GetContentEnglish & IIf(bolCallMail = True, "   ", "") & "   Tai E Alert Date: " & ChgEngDate(np(0)) & vbCrLf
                     'end 2021/8/2
                  Else
                     'Modified by Morgan 2021/8/2
                     'GetContentEnglish = GetContentEnglish & IIf(bolCallMail = True, "   ", "") & "   Deadline: " & vbCrLf
                     GetContentEnglish = GetContentEnglish & IIf(bolCallMail = True, "   ", "") & "   Tai E Alert Date: " & vbCrLf
                     'end 2021/8/2
                  End If
               End If
            
            End If
         End If
         
      Case "稱謂"
'Modified by Morgan 2024/4/10 對外統一用 Dear Colleagues --林總
'         '語文為英文並且為FC代理人名稱
'         'Modify By Sindy 2014/9/25
'         If bolCallMail = True Then
'            'Added by Morgan 2024/3/15 客戶要求--Lisa
'            'Modified by Morgan 2024/3/27
'            'If ChangeCustomerL(strExc(0)) = "Y22327000" Then
'            strExc(0) = ""
'            If pa(1) = "FCP" Or pa(1) = "P" Or pa(1) = "CFP" Then
'               strExc(0) = pa(75)
'            ElseIf pa(1) = "FG" Or pa(1) = "PS" Or pa(1) = "CPS" Then
'               strExc(0) = pa(26)
'            End If
'            If ChangeCustomerL(strExc(0)) = "Y22327000" Then
'            'end 2024/3/27
'               GetContentEnglish = "Dear Colleagues," & vbCrLf
'            Else
'            'end 2024/3/15
'               GetContentEnglish = "Dear Sirs," & vbCrLf
'            End If
'         Else
'         '2014/9/25 END
'            If Option1(1).Value = True And Option2.Value = True Then
'               'Added by Morgan 2021/2/2 寶齡富錦 Y55435 特殊控制
'               If bIsBPFCase Then
'                  GetContentEnglish = "Dear Associates," & vbCrLf
'               Else
'               'end 2021/2/2
'                  'Added by Morgan 2024/3/15 客戶要求--Lisa
'                  'Modified by Morgan 2024/3/27
'                  'If ChangeCustomerL(strExc(0)) = "Y22327000" Then
'                  strExc(0) = ""
'                  If pa(1) = "FCP" Or pa(1) = "P" Or pa(1) = "CFP" Then
'                     strExc(0) = pa(75)
'                  ElseIf pa(1) = "FG" Or pa(1) = "PS" Or pa(1) = "CPS" Then
'                     strExc(0) = pa(26)
'                  End If
'                  If ChangeCustomerL(strExc(0)) = "Y22327000" Then
'                  'end 2024/3/27
'                     GetContentEnglish = "Dear Colleagues," & vbCrLf
'                  Else
'                  'end 2024/3/15
'                     GetContentEnglish = "Dear Sirs," & vbCrLf
'
'                  End If
'               End If
'            Else
'               GetContentEnglish = "Dear Associates," & vbCrLf
'            End If
'         End If
         GetContentEnglish = "Dear Colleagues," & vbCrLf
'end 2024/4/10
         
'      Case "敬語"
'         If Left(m_Dept, 2) = "F2" Or pa(1) = "FCP" Or pa(1) = "FG" Then
'            .Selection.TypeParagraph
'            .Selection.TypeParagraph
'            .Selection.TypeText Space(30) & "Best regards,"
'            .Selection.TypeParagraph
'            .Selection.TypeParagraph
'            .Selection.TypeParagraph
'            .Selection.TypeParagraph
'            .Selection.TypeText Space(30) & "Fred C. T. Yen"
'            .Selection.TypeParagraph
'            .Selection.TypeText "Patent Department" & Space(15) & "Patent Attorney"
'            .Selection.TypeParagraph
'            If m_CompNo = "J" Then
'               .Selection.TypeText Space(30) & "Tai E Intellectual Property Co., Ltd."
'            Else
'               .Selection.TypeText Space(30) & "Tai E International Patent & Law Office"
'            End If
'            .Selection.TypeParagraph
'            .Selection.TypeText "CTY/" & Pub_StrUserSt17
'
'         ElseIf pa(1) = "CFP" Or pa(1) = "P" Then
'            .Selection.TypeParagraph
'            .Selection.TypeParagraph
'            .Selection.ParagraphFormat.FirstLineIndent = .CentimetersToPoints(7.35)
'            .Selection.TypeText "Best regards,"
'            .Selection.TypeParagraph
'            .Selection.TypeParagraph
'            .Selection.TypeParagraph
'            .Selection.TypeParagraph
'            .Selection.TypeText "Jerry C. Y. Lin"
'            .Selection.TypeParagraph
'            .Selection.TypeText "Patent Attorney"
'            .Selection.TypeParagraph
'            .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
'            .Selection.TypeParagraph
'            .Selection.ParagraphFormat.FirstLineIndent = 0
'            .Selection.TypeText "CYL/" & Pub_StrUserSt17
'
'         Else
'            .Selection.TypeParagraph
'            .Selection.TypeParagraph
'            .Selection.TypeText Space(30) & "Best regards,"
'            .Selection.TypeParagraph
'            .Selection.TypeParagraph
'            .Selection.TypeParagraph
'            If m_CompNo = "J" Then
'               .Selection.TypeText Space(30) & "Tai E Intellectual"
'            Else
'               .Selection.TypeText Space(30) & "Tai E International"
'            End If
'            .Selection.TypeParagraph
'            If m_CompNo = "J" Then
'               .Selection.TypeText Space(30) & "Property Co., Ltd."
'            Else
'               .Selection.TypeText Space(30) & "Patent & Law Office"
'            End If
'            .Selection.TypeParagraph
'            .Selection.TypeText "CTY/" & Pub_StrUserSt17
'         End If
   End Select
   stCaseNo = Trim(stCaseNo)
   stYourRef = Trim(stYourRef)
   stOurRef = Trim(stOurRef)
End Function

Private Sub Form_Initialize()
   'add by nickc 2006/07/12
   ReDim pa(TF_PA) As String
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me

   lblClose.Caption = ""
   '限制智權人員不能選代理人
   'Modify By Sindy 2014/10/14 杜經理可選代理人
   'If Left(Pub_StrUserSt15, 1) = "S" Then
   If Left(Pub_StrUserSt03, 1) = "S" Then
      bolFA = False
   Else
      bolFA = True
   End If
   Option2.Enabled = bolFA
   Option6.Enabled = bolFA
   
   'ADD BY SONIA 2015/10/16 專利處人員不顯示代理人備註及申請人備註
   If Left(Pub_StrUserSt03, 2) = "P1" Then
      Label7.Visible = False
      Text6(0).Visible = False
      Label8.Visible = False
      Text6(1).Visible = False
   Else
      Label7.Visible = True
      Text6(0).Visible = True
      Label8.Visible = True
      Text6(1).Visible = True
   End If
   'END 2015/10/16
   
   'Add By Sindy 2014/8/15
   'Modify By Sindy 2014/9/12
   If Pub_StrUserSt03 = "M51" Then
      'Modify By Sindy 2014/10/8
      If bolFrom1105Callme = True Then
         m_StrUserST03 = "F22"
      Else
      '2014/10/8 END
         m_StrUserST03 = UCase(InputBox("請輸入欲操作的部門代號？" & vbCrLf & "(F22:外專程序 F23:外專承辦 P22:商標處程序)"))
      End If
   End If
   If m_StrUserST03 = "" Then m_StrUserST03 = Pub_StrUserSt03
   'Added by Lydia 2020/06/12 各項指示之使用部門
   If m_StrUserST03 <> "" Then
       Select Case Left(m_StrUserST03, 2)
            Case "F1", "P1"
                  '商標
                  m_strIT10 = "T"
            Case "F2", "P2"
                  '專利
                  m_strIT10 = "P"
            Case Else
                  m_strIT10 = "T"
       End Select
       If m_StrUserST03 = "M51" Then m_strIT10 = "P,T"
   End If
   'end 2020/06/20
   'Added by Lydiap 2020/09/11 增加「各項指示」勾選項，勾選後在執行FC郵件一併另外開啟撰寫信函Word檔
   ChkINST.Visible = False
   ChkINST.Value = 0
   If InStr("F1,F2", Left(m_StrUserST03, 2)) > 0 Then
       ChkINST.Visible = True
       ChkINST.Value = 1  '(預設勾選)
   End If
   'end 2020/09/11
   
   'Added by Lydia 2021/02/08 增加「(各項指示)含欄位設定」勾選項，勾選後才會抓各項指示的欄位設定記錄(by David)
   ChkINSTdef.Top = ChkINST.Top
   ChkINSTdef.Visible = False
   ChkINSTdef.Value = 0   '預設：不含欄位設定
   If InStr("F1,F2", Left(m_StrUserST03, 2)) > 0 Then
       ChkINSTdef.Visible = True
   End If
   'end 2021/02/08
   
   cmdFCMail(1).Visible = False
   If m_StrUserST03 = "M51" Or m_StrUserST03 = "F22" Then
      cmdFCMail(0).Caption = "發FC郵件(工程師署名)"
      cmdFCMail(1).Visible = True
      'Download郵件範本
      strTemplatePath = PUB_DownloadOftPath("F23", "", EMailType)
   'Add By Sindy 2017/3/6
   ElseIf m_StrUserST03 = "P22" Then
      cmdFCMail(0).Caption = "巨京查名郵件"
      'Download郵件範本
      strTemplatePath = PUB_DownloadOftPath("P22", "P29", EMailType)
   '2017/3/6 END
   Else
      cmdFCMail(0).Caption = "發 FC 郵件"
   End If
   
   'Modify By Sindy 2014/9/18
   'Download郵件範本
   strTemplatePath = PUB_DownloadOftPath(m_StrUserST03, Pub_StrUserSt17, EMailType)

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Call ChkTransFile 'Added by Lydia 2017/06/16
   
   Set frm090401 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
   
   'Added by Morgan 2011/12/7 改 T 案對代理人的中文信預設不要信頭其餘都要
   'modify by sonia 2014/4/28
   'If Text1.Text = "T" And Option1(0).Value = True And Option3.Value = False Then
   If InStr(Text1.Text, "T") > 0 And Option1(0).Value = True And Option3.Value = False Then
      txtLetterHead = "N"
   Else
      txtLetterHead = ""
   End If
   'end 2011/12/7
      
   'Modify by Morgan 2007/3/5 加日文
   'If Option1(1).Value = True Then
   If Option1(1).Value = True Or Option1(2).Value = True Then
      
'Removed by Morgan 2011/12/7 改 T 案預設不要信頭其餘都要(控制移到上面)
'      'Modify By Sindy 2009/07/24 增加LIN系統類別
'      If (Text1.Text = "FCP" Or Text1.Text = "FG" Or Text1.Text = "CFT" Or _
'            Text1.Text = "FCT" Or Text1.Text = "FCL" Or Text1.Text = "CFL" Or Text1.Text = "LIN") Then
'         txtLetterHead = ""
'      'Add by Morgan 2007/3/14 外專寫大陸案時
'      ElseIf (Text1.Text = "P" And index = 1) Then
'         txtLetterHead = ""
'
'      'Add by Morgan 2011/7/12 CFP改預設要印信頭
'      ElseIf Text1 = "CFP" Then
'         txtLetterHead = ""
'
'      Else
'         txtLetterHead = "N"
'      End If
'end 2011/12/7

      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/30 +ACS系統類別
      If (Text1.Text = "CFP" Or Text1.Text = "CPS" Or Text1.Text = "FCP" Or _
          Text1.Text = "FG" Or Text1.Text = "CFT" Or Text1.Text = "FCT" Or _
          Text1.Text = "FCL" Or Text1.Text = "CFL" Or Text1.Text = "LIN" Or Text1.Text = "ACS") Then
         'Modified by Morgan 2018/10/23 預設不印傳真封面--慧汶
         'txtFaxFace.Text = ""
         If Text1.Text = "CFP" Or Text1.Text = "CPS" Then
            txtFaxFace.Text = "N"
         Else
            txtFaxFace.Text = ""
         End If
         'end 2018/10/23
         txtFaxFace.Enabled = True
      'Add by Morgan 2007/3/14 外專寫大陸案時
      ElseIf (Text1.Text = "P" And Index = 1) Then
         txtFaxFace.Text = ""
         txtFaxFace.Enabled = True
      Else
         txtFaxFace.Text = "N"
         txtFaxFace.Enabled = False
      End If

   Else
      txtFaxFace.Text = "N"
      txtFaxFace.Enabled = False
      
'Removed by Morgan 2011/12/7 改 T 案預設不要信頭其餘都要(控制移到上面)
'      'Modify by Morgan 2010/12/13 中文信也可印信頭,業務預設要其餘否
'      'txtLetterHead.Text = "N"
'      'Modify by Morgan 2011/6/1 改判斷選申請人,中文信時預設要印信頭
'      'If Left(Pub_StrUserSt15, 1) = "S" Then
'      If Option3.Value = True Then
'         txtLetterHead.Text = ""
'      Else
'         txtLetterHead.Text = "N"
'      End If
'      'end 2010/12/13
'end 2011/12/7
      
   End If
   
   'Add by Morgan 2007/3/12
   If pa(1) = "FCP" Then
      If m_iLang <> -1 And Index <> m_iLang Then
         MsgBox "信函語文已變更(與預設不同)！", vbExclamation
      End If
   End If
   
   'add by sonia 2018/10/17 有FC代理人且案件之智權人員非國外部時選中文,自動預設發信對象為FC代理人
   If Option1(0).Value = True And Left(GetSalesArea(PUB_GetAKindSalesNo(Me.Text1.Text, Me.Text2.Text, Me.Text3.Text, Me.Text4.Text)), 1) <> "F" Then
      Select Case pa(1)
         Case "P", "CFP", "FCP"        '專利
            If pa(75) <> "" Then
               Option2.Value = True
            End If
         Case "T", "TF", "CFT", "FCT"  '商標
            If pa(44) <> "" Then
               Option2.Value = True
            End If
         'modify by sonia 2019/7/30 +ACS系統類別
         Case "L", "FCL", "CFL", "LIN", "ACS" '法務
            If pa(22) <> "" Then
               Option2.Value = True
            End If
         Case Else                     '服務業務
            If pa(26) <> "" And pa(1) <> "LA" Then
               Option2.Value = True
            End If
      End Select
   End If
   'end 2018/10/17
End Sub

'點選FC代理人
Private Sub Option2_Click()
   If CaseNoCheck = False Then Exit Sub
   
   'Add by Morgan 2010/12/22
   '中文信預設不要信頭 --郭
   'Modified by Morgan 2011/12/7 改 T 案對代理人的中文信預設不要信頭其餘都要
   'If Option1(0).Value Then
   'modify by sonia 2014/4/28
   'If Text1.Text = "T" And Option1(0).Value = True And Option3.Value = False Then
   If InStr(Text1.Text, "T") > 0 And Option1(0).Value = True And Option3.Value = False Then
      txtLetterHead = "N"
   Else
   'end 2010/12/22
      txtLetterHead = "" 'Add by Morgan 2006/11/28 FC代理人預設印信頭
   End If
   
   Select Case Text1.Text
      Case "P", "CFP", "FCP"
          m_strFCAgent = pa(75)
          m_strContact1(0) = pa(51): m_strContact1(1) = pa(52): m_strContact1(2) = pa(53)
          m_strContact2(0) = pa(54): m_strContact2(1) = pa(55): m_strContact2(2) = pa(56)
          m_strContact3(0) = pa(98): m_strContact3(1) = pa(99): m_strContact3(2) = pa(100)
          m_strConDepJp = pa(139) 'Add by Morgan 2007/3/6
      Case "T", "TF", "CFT", "FCT"
          m_strFCAgent = pa(44)
          m_strContact1(0) = pa(38): m_strContact1(1) = pa(39): m_strContact1(2) = pa(40)
          m_strContact2(0) = pa(41): m_strContact2(1) = pa(42): m_strContact2(2) = pa(43)
          m_strContact3(0) = "": m_strContact3(1) = "": m_strContact3(2) = ""
          m_strConDepJp = pa(76) 'Add by Morgan 2007/3/6
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/30 +ACS系統類別
      Case "L", "FCL", "CFL", "LIN", "ACS"
          m_strFCAgent = pa(22)
          m_strContact1(0) = pa(18): m_strContact1(1) = pa(19): m_strContact1(2) = pa(20)
          m_strContact2(0) = "": m_strContact2(1) = "": m_strContact2(2) = ""
          m_strContact3(0) = "": m_strContact3(1) = "": m_strContact3(2) = ""
          m_strConDepJp = pa(39) 'Add by Morgan 2007/3/6
      Case "LA"
          m_strFCAgent = ""
          m_strContact1(0) = "": m_strContact1(1) = "": m_strContact1(2) = ""
          m_strContact2(0) = "": m_strContact2(1) = "": m_strContact2(2) = ""
          m_strContact3(0) = "": m_strContact3(1) = "": m_strContact3(2) = ""
      Case Else
          m_strFCAgent = pa(26)
          m_strContact1(0) = pa(30): m_strContact1(1) = "": m_strContact1(2) = ""
          m_strContact2(0) = "": m_strContact2(1) = "": m_strContact2(2) = ""
          m_strContact3(0) = "": m_strContact3(1) = "": m_strContact3(2) = ""
          m_strConDepJp = pa(71) 'Add by Morgan 2007/3/6
   End Select
   
   '若有FC代理人
   If m_strFCAgent <> "" Then
      'Modify by Morgan 2007/1/24 加 FA70,FA32,FA33,FA34,FA35,FA36
      'Modify by Morgan 2007/3/6 加 FA78
      strExc(0) = "SELECT FA04,FA05,FA63,FA64,FA65,FA06,FA07,FA08,FA09,FA52,FA53,FA54" & _
         ",FA29,FA56,FA57,FA58,FA17,FA18,FA19,FA20,FA21,FA22,FA70,FA23,FA32,FA33,FA34,FA35,FA36" & _
         ",FA12,FA14,FA13,FA15,FA78 FROM FAGENT WHERE " & ChgFagent(m_strFCAgent)
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         '若基本檔有聯絡人1
         If Trim(m_strContact1(0) & m_strContact1(1) & m_strContact1(2)) <> "" Then
            If m_strContact1(0) <> "" Then Combo4.AddItem m_strContact1(0)
            If m_strContact1(1) <> "" Then Combo4.AddItem m_strContact1(1)
            If m_strContact1(2) <> "" Then Combo4.AddItem m_strContact1(2)
            Combo4 = m_strContact1(0)
            txtConDepJP.Text = m_strConDepJp 'Add by Morgan 2007/3/6
         Else
            'FC聯絡人1
            With RsTemp
               For i = 6 To 8
                  If IsNull(.Fields(i)) = False And (.Fields(i)) <> "" Then
                      fa(i) = .Fields(i)
                      Combo4.AddItem fa(i)
                      Combo4 = fa(6)
                  End If
               Next
               txtConDepJP.Text = "" & .Fields("FA78") 'Add by Morgan 2007/3/6
          
            End With
         End If
         '若基本檔有聯絡人2
         If Trim(m_strContact2(0) & m_strContact2(1) & m_strContact2(2)) <> "" Then
            If m_strContact2(0) <> "" Then Combo5.AddItem m_strContact2(0)
            If m_strContact2(1) <> "" Then Combo5.AddItem m_strContact2(1)
            If m_strContact2(2) <> "" Then Combo5.AddItem m_strContact2(2)
            Combo5 = m_strContact2(0)
         Else
            'FC聯絡人2
            With RsTemp
                For i = 9 To 11
                    If IsNull(.Fields(i)) = False And (.Fields(i)) <> "" Then
                        fa(i) = .Fields(i)
                        Combo5.AddItem fa(i)
                        Combo5 = fa(9)
                    End If
                Next
            End With
         End If
         
         '若基本檔有實體聯絡人
         If Trim(m_strContact3(0) & m_strContact3(1) & m_strContact3(2)) <> "" Then
             If m_strContact3(0) <> "" Then Combo6.AddItem m_strContact3(0)
             If m_strContact3(1) <> "" Then Combo6.AddItem m_strContact3(1)
             If m_strContact3(2) <> "" Then Combo6.AddItem m_strContact3(2)
             Combo6 = m_strContact3(0)
         Else
             'FC實體聯絡人
             With RsTemp
                 For i = 13 To 15
                     If IsNull(.Fields(i)) = False And (.Fields(i)) <> "" Then
                         fa(i) = .Fields(i)
                         Combo6.AddItem fa(i)
                         Combo6 = fa(13)
                     End If
                 Next
             End With
         End If
         
         If Combo4.ListCount > 0 And Combo4.ListIndex = -1 Then Combo4.ListIndex = 0
         If Combo5.ListCount > 0 And Combo5.ListIndex = -1 Then Combo5.ListIndex = 0
         If Combo6.ListCount > 0 And Combo6.ListIndex = -1 Then Combo6.ListIndex = 0
         
         With RsTemp
            'FC代理人備註
            '2006/12/11 MODIFY BY SONIA 改與 OPTION6 寫法相同
            'If IsNull(.Fields(12)) = False And (.Fields(12)) = "" Then fa(12) = .Fields(12)
            If Not IsNull(.Fields(12)) Then fa(12) = .Fields(12)
            Text6(0) = fa(12)
       
            'Modify by Morgan 2007/1/24 改先抓POB, 加英文地址6
            'FC代理人地址
            '中
            fa(16) = "" & .Fields("FA17")
            '英
            'POB
            If Not IsNull(.Fields("FA32")) Then
               For i = 17 To 21
                  fa(i) = "" & .Fields(i + 7)
               Next
               fa(22) = ""
            '地址
            Else
               For i = 17 To 22
                  If Not IsNull(.Fields(i)) Then fa(i) = .Fields(i)
               Next
            End If
            '日
            fa(23) = "" & .Fields("FA23")
       
            'Add by Morgan 2006/2/13
            m_strFax(1) = "" & .Fields("FA12")
            m_strFax(0) = "" & .Fields("FA14")
            'Add by Morgan 2007/1/19
            m_strFax(3) = "" & .Fields("FA13")
            m_strFax(2) = "" & .Fields("FA15")
         End With
         
         'Modified by Lydia 2020/06/04 備註和各項指示並存(暫時到完全取代), 原本是抓各項指示取代客戶備註
         strMemoY = "": strKeyY = ""
         If strSrvDate(1) >= 各項指示啟用日 Then
             'Added by Lydia 2020/08/26 各項指示：完成確認，各項指示取代原先備註
             If PUB_GetInstConfirm(m_StrUserST03, m_strFCAgent) = True Then
                  Text6(0) = ""
             End If
             'end 2020/08/26
             If Pub_GetInstructions(Me.Name, m_strFCAgent, strMemoY, , , , m_strIT10) Then
                 If strMemoY <> "" Then
                     Text6(0) = Text6(0) & IIf(Text6(0) <> "", vbCrLf, "") & strMemoY
                     strKeyY = m_strFCAgent
                 End If
             End If
         End If
         'end 2020/06/04
      End If
   End If
End Sub

Private Sub Option3_Click()
   Dim stContNo As String
   
   If CaseNoCheck = False Then Exit Sub
   
   'Modified by Morgan 2011/12/7 改 T 案對代理人的中文信預設不要信頭其餘都要
   'txtLetterHead = "N" 'Add by Morgan 2006/11/28  申請人預設不印信頭
   'Add by Morgan 2010/12/14 業務預設要印信頭
   'Modify by Morgan 2011/6/1 改判斷選申請人,中文信時預設要印信頭
   'If Left(Pub_StrUserSt15, 1) = "S" Then
   'If Option1(0).Value = True Then
   '   txtLetterHead.Text = ""
   'End If
   'modify by sonia 2014/4/28
   'If Text1.Text = "T" And Option1(0).Value = True And Option3.Value = False Then
   If InStr(Text1.Text, "T") > 0 And Option1(0).Value = True And Option3.Value = False Then
      txtLetterHead = "N"
   Else
      txtLetterHead = ""
   End If
   'end 2011/12/7
   
'Modified by Morgan 2014/9/12 申請人不可抓基本檔聯絡人
    Select Case Text1.Text
    Case "P", "CFP", "FCP"
        m_CustNo(1) = pa(26)
'        m_strContact1(0) = pa(51): m_strContact1(1) = pa(52): m_strContact1(2) = pa(53)
'        m_strContact2(0) = pa(54): m_strContact2(1) = pa(55): m_strContact2(2) = pa(56)
'        m_strContact3(0) = pa(98): m_strContact3(1) = pa(99): m_strContact3(2) = pa(100)
    Case "T", "TF", "CFT", "FCT"
        m_CustNo(1) = pa(23)
'        m_strContact1(0) = pa(38): m_strContact1(1) = pa(39): m_strContact1(2) = pa(40)
'        m_strContact2(0) = pa(41): m_strContact2(1) = pa(42): m_strContact2(2) = pa(43)
'        m_strContact3(0) = "": m_strContact3(1) = "": m_strContact3(2) = ""
    Case "L"
        m_CustNo(1) = pa(11)
'        m_strContact1(0) = pa(18): m_strContact1(1) = pa(19): m_strContact1(2) = pa(20)
'        m_strContact2(0) = "": m_strContact2(1) = "": m_strContact2(2) = ""
'        m_strContact3(0) = "": m_strContact3(1) = "": m_strContact3(2) = ""
    Case "LA"
        m_CustNo(1) = pa(5)
'        m_strContact1(0) = "": m_strContact1(1) = "": m_strContact1(2) = ""
'        m_strContact2(0) = "": m_strContact2(1) = "": m_strContact2(2) = ""
'        m_strContact3(0) = "": m_strContact3(1) = "": m_strContact3(2) = ""
    Case Else
        m_CustNo(1) = pa(8)
'        m_strContact1(0) = pa(30): m_strContact1(1) = "": m_strContact1(2) = ""
'        m_strContact2(0) = "": m_strContact2(1) = "": m_strContact2(2) = ""
'        m_strContact3(0) = "": m_strContact3(1) = "": m_strContact3(2) = ""
    End Select
      Erase m_strContact1
      Erase m_strContact2
      Erase m_strContact3
      m_strConDepJp = ""
      Combo4.Clear
      Combo4 = "" 'Added by Morgan 2015/10/8 combo 可能會殘留上次顯示
      Combo5.Clear
      Combo5 = "" 'Added by Morgan 2015/10/8 combo 可能會殘留上次顯示
      Combo6.Clear
      Combo6 = "" 'Added by Morgan 2015/10/8 combo 可能會殘留上次顯示
      txtConDepJP.Text = m_strConDepJp
'end 2014/9/12

    '若點選申請人
    If Option3.Value = True Then
        If m_CustNo(1) <> "" Then
            '抓客戶基本檔
            '中文地址 : 聯絡地址-->中文申請地址
            'Modify by Morgan 2007/1/24 加CU102,CU65,CU66,CU67,CU68,CU69
            'Modify by Morgan 2008/7/17 郵遞區號分開抓
            'Modify by Morgan 2008/8/1 接洽人改先抓個案，若個案沒有再用客戶聯絡人編號抓聯絡人檔,若有聯絡人編號但該聯絡人無地址時抓原客戶地址
            'Modified by Morgan 2023/11/29 +TEL/FAX改先抓接洽人 CU16,CU18,CU17,CU19->NVL(PCC30,CU16) CU16,NVL(PCC31,CU18) CU18,DECODE(PCC30,NULL,CU17) CU17,DECODE(PCC31,NULL,CU19) CU19
            stContNo = GetCaseContactNo(Text1, Text2, Text3, Text4, m_CustNo(1))
            strExc(0) = "SELECT CU04,CU05,CU88,CU89,CU90,CU06,CU58,CU59,CU60,CU61,CU62,CU63" & _
               ",CU79,CU91,CU92,CU93, DECODE(PCC22,NULL,Decode(CU31, Null, CU23,CU31),PCC22),CU24,CU25,CU26,CU27,CU28,CU102,CU29" & _
               ",CU65,CU66,CU67,CU68,CU69,NVL(PCC30,CU16) CU16,NVL(PCC31,CU18) CU18,DECODE(PCC30,NULL,CU17) CU17,DECODE(PCC31,NULL,CU19) CU19" & _
               ",NVL(PCC21,CU30) CU30,NVL(PCC05,CU08) CU08,CU80,CU104 FROM CUSTOMER,POTCUSTCONT WHERE " & ChgCustomer(m_CustNo(1)) & " AND PCC01(+)=CU01 "
            If stContNo = "" Then
               strExc(0) = strExc(0) & " AND PCC02(+)=CU127 "
            Else
               strExc(0) = strExc(0) & " AND PCC02(+)='" & stContNo & "'"
            End If
            'end 2008/8/1
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                '若基本檔有聯絡人
                If m_strContact1(0) & m_strContact1(1) & m_strContact1(2) <> "" Then
                    If m_strContact1(0) <> "" Then Combo4.AddItem m_strContact1(0)
                    If m_strContact1(1) <> "" Then Combo4.AddItem m_strContact1(1)
                    If m_strContact1(2) <> "" Then Combo4.AddItem m_strContact1(2)
                    Combo4 = m_strContact1(0)
                Else
                    '申請人聯絡人1
                    With RsTemp
                        For i = 6 To 8
                            If IsNull(.Fields(i)) = False And (.Fields(i)) <> "" Then
                                cu(i) = .Fields(i)
                                Combo4.AddItem cu(i)
                                Combo4 = cu(4)
                            End If
                        Next
                    End With
                End If
          
                '若基本檔有聯絡人
                If m_strContact1(0) & m_strContact1(1) & m_strContact1(2) <> "" Then
                    If m_strContact2(0) <> "" Then Combo5.AddItem m_strContact2(0)
                    If m_strContact2(1) <> "" Then Combo5.AddItem m_strContact2(1)
                    If m_strContact2(2) <> "" Then Combo5.AddItem m_strContact2(2)
                    Combo5 = m_strContact3(0)
                Else
                    '申請人聯絡人2
                    With RsTemp
                        For i = 9 To 11
                            If IsNull(.Fields(i)) = False And (.Fields(i)) <> "" Then
                                cu(i) = .Fields(i)
                                Combo5.AddItem cu(i)
                                Combo5 = cu(9)
                            End If
                        Next
                    End With
                End If
          
                '若基本檔有實體聯絡人
                If m_strContact3(0) & m_strContact3(1) & m_strContact3(2) <> "" Then
                    If m_strContact3(0) <> "" Then Combo6.AddItem m_strContact3(0)
                    If m_strContact3(1) <> "" Then Combo6.AddItem m_strContact3(1)
                    If m_strContact3(2) <> "" Then Combo6.AddItem m_strContact3(2)
                    Combo6 = m_strContact3(0)
                Else
                    '申請人實體聯絡人
                    With RsTemp
                        For i = 13 To 15
                            If IsNull(.Fields(i)) = False And (.Fields(i)) <> "" Then
                                cu(i) = .Fields(i)
                                Combo6.AddItem cu(i)
                                Combo6 = cu(13)
                            End If
                        Next
                    End With
                End If
          
                If Combo4.ListCount > 0 And Combo4.ListIndex = -1 Then Combo4.ListIndex = 0
                If Combo5.ListCount > 0 And Combo5.ListIndex = -1 Then Combo5.ListIndex = 0
                If Combo6.ListCount > 0 And Combo6.ListIndex = -1 Then Combo6.ListIndex = 0
           
                With RsTemp
                  '申請人備註
                  If Not IsNull(.Fields(12)) Then cu(12) = .Fields(12)
                  Text6(1) = cu(12)
                  '申請人地址
                  'Modify by Morgan 2007/1/24 改先抓POB, 加英文地址6
                  '中
                  cu(16) = "" & .Fields(16)
                  'Add by Morgan 2008/7/17
                  m_Zip = "" & .Fields("cu30")
                  m_Contact = "" & .Fields("cu08")
                  m_CU80 = "" & .Fields("cu80")
                  'end 2008/7/17
                  m_CU104 = "" & .Fields("cu104") 'Add by Morgan 2008/8/6
                  '英
                  'POB
                  If Not IsNull(.Fields("CU65")) Then
                     For i = 17 To 21
                        cu(i) = "" & .Fields(i + 7)
                     Next
                     cu(22) = ""
                  '地址
                  Else
                     For i = 17 To 22
                        cu(i) = "" & .Fields(i)
                     Next
                  End If
                  '日
                  cu(23) = "" & .Fields("CU29")
                  'end 2007/1/24
       
                  'Add by Morgan 2006/2/13
                  m_strFax(1) = "" & .Fields("CU16")
                  m_strFax(0) = "" & .Fields("CU18")
                  'Add by Morgan 2007/1/19
                  m_strFax(3) = "" & .Fields("CU17")
                  m_strFax(2) = "" & .Fields("CU19")
                End With
                'Modified by Lydia 2020/06/04 備註和各項指示並存(暫時到完全取代), 原本是抓各項指示取代客戶備註
                strMemoX = "": strKeyX = ""
                If strSrvDate(1) >= 各項指示啟用日 Then
                    'Added by Lydia 2020/08/26 各項指示：完成確認，各項指示取代原先備註
                    If PUB_GetInstConfirm(m_StrUserST03, m_CustNo(1)) = True Then
                         Text6(1) = ""
                    End If
                    'end 2020/08/26
                    If Pub_GetInstructions(Me.Name, m_CustNo(1), strMemoX, , , , m_strIT10) Then
                        If strMemoX <> "" Then
                            Text6(1) = Text6(1) & IIf(Text6(1) <> "", vbCrLf, "") & strMemoX
                            strKeyX = m_CustNo(1)
                        End If
                    End If
                End If
                'end 2020/06/04
            End If
        End If
    End If
End Sub

'點選CF代理人
Private Sub Option6_Click()
   If CaseNoCheck = False Then Exit Sub
   
   'Modified by Morgan 2011/12/7 改 T 案對代理人的中文信預設不要信頭其餘都要
   ''Add by Morgan 2011/7/4 CFP改預設要印信頭
   'If Text1 = "CFP" Then
   '   txtLetterHead = ""
   'Else
   ''end 2011/7/4
   '   txtLetterHead = "N" 'Add by Morgan 2006/11/28 CF代理人預設不印信頭
   'End If 'Add by Morgan 2011/7/4
   'modify by sonia 2014/4/28
   'If Text1.Text = "T" And Option1(0).Value = True And Option3.Value = False Then
   If InStr(Text1.Text, "T") > 0 And Option1(0).Value = True And Option3.Value = False Then
      txtLetterHead = "N"
   Else
      txtLetterHead = ""
   End If
   'end 2011/12/7
   
   
   '以本所案號抓案件進度檔發文日非0且日期最大的cp44代理人
   '2006/5/23 MODIFY BY SONIA 加 CP09<'C'條件
   'Modify by Morgan 2008/2/18 改以發文日排序
   '2009/6/23 MODIFY BY SONIA 郭雅娟說取消發文日CP27>0限制,因大陸已輸指示信未發文前無法使用
   'Modify by Morgan 2010/9/7 +CP45,CP09
   'Modify by Morgan 2011/1/5 +CP57 is null
   'Added by Lydia 2019/02/26 CFP之EPC子案於撰寫信函抓指定代理人(與期限通知管制表及年費發文指示信之子案代理人抓法相同)
   If m_ChildCP44 <> "" Then
        If PUB_GetEPCtoCP44(pa(1), pa(2), pa(3), pa(4), m_strCP44, strExc(1), m_strCP45, m_strCP09) = True Then
        End If
   'Added by Morgan 2023/12/13 CFP美國發明案CF代理人特殊規則
   'Modified by Morgan 2024/3/27 CFP案非美國發明案應該可以不必重抓(案號輸入後已有用函數抓過)
   'ElseIf pa(1) = "CFP" And pa(9) = "101" And pa(8) = "1" Then
   ElseIf pa(1) = "CFP" Then
      If pa(9) = "101" And pa(8) = "1" Then
   'end 2024/3/27
         SetCFPUSAgent
      End If
   Else
   'end 2019/02/26
   
      m_strCP44 = ""
      strExc(0) = "select CP44,CP45,CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
         " AND CP44 Is Not Null AND CP09<'C' and cp57 is null"
      'Modify By Sindy 2014/9/4 CFT撰寫信函之CF代理人名稱：文件簽證711及申請英文證明304 不要列入
      If pa(1) = "CFT" Then
        strExc(0) = strExc(0) & " AND CP10 NOT IN ('711','304')"
      End If
      
      'Added by Morgan 2024/4/11 排除P大陸案年費--郭
      If pa(1) = "P" And pa(9) = "020" Then
         strExc(0) = strExc(0) & " AND CP10<>'605'"
      End If
      'end 2024/4/11
      
      strExc(0) = strExc(0) & " Order By CP27 Desc, CP09 Desc"
      '2014/9/4 End
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_strCP44 = "" & RsTemp.Fields(0).Value
         'Add by Morgan 2010/9/7
         m_strCP45 = "" & RsTemp.Fields(1).Value
         m_strCP09 = "" & RsTemp.Fields(2).Value
      End If
   End If 'end 2019/02/26
   
   If m_strCP44 <> "" Then
      If Combo7.Tag <> m_strCP44 Then SetCFRef 'Added by Morgan 2023/12/13
      
'Modified by Morgan 2014/9/12 CF代理人不可抓基本檔聯絡人
'      Select Case Text1.Text
'         Case "P", "CFP", "FCP"
'            m_strContact1(0) = pa(51): m_strContact1(1) = pa(52): m_strContact1(2) = pa(53)
'            m_strContact2(0) = pa(54): m_strContact2(1) = pa(55): m_strContact2(2) = pa(56)
'            m_strContact3(0) = pa(98): m_strContact3(1) = pa(99): m_strContact3(2) = pa(100)
'         Case "T", "TF", "CFT", "FCT"
'            m_strContact1(0) = pa(38): m_strContact1(1) = pa(39): m_strContact1(2) = pa(40)
'            m_strContact2(0) = pa(41): m_strContact2(1) = pa(42): m_strContact2(2) = pa(43)
'            m_strContact3(0) = "": m_strContact3(1) = "": m_strContact3(2) = ""
'         Case "L"
'            m_strContact1(0) = pa(18): m_strContact1(1) = pa(19): m_strContact1(2) = pa(20)
'            m_strContact2(0) = "": m_strContact2(1) = "": m_strContact2(2) = ""
'            m_strContact3(0) = "": m_strContact3(1) = "": m_strContact3(2) = ""
'         Case "LA"
'            m_strContact1(0) = "": m_strContact1(1) = "": m_strContact1(2) = ""
'            m_strContact2(0) = "": m_strContact2(1) = "": m_strContact2(2) = ""
'            m_strContact3(0) = "": m_strContact3(1) = "": m_strContact3(2) = ""
'         Case Else
'            m_strContact1(0) = pa(30): m_strContact1(1) = "": m_strContact1(2) = ""
'            m_strContact2(0) = "": m_strContact2(1) = "": m_strContact2(2) = ""
'            m_strContact3(0) = "": m_strContact3(1) = "": m_strContact3(2) = ""
'      End Select
            Erase m_strContact1
            Erase m_strContact2
            Erase m_strContact3
            m_strConDepJp = ""
            Combo4.Clear
            Combo4 = "" 'Added by Morgan 2015/10/8 combo 可能會殘留上次顯示
            Combo5.Clear
            Combo5 = "" 'Added by Morgan 2015/10/8 combo 可能會殘留上次顯示
            Combo6.Clear
            Combo6 = "" 'Added by Morgan 2015/10/8 combo 可能會殘留上次顯示
            txtConDepJP.Text = m_strConDepJp
'end 2014/9/12
         '抓CFAgent
         'Modify by Morgan 2007/1/24 加FA70,FA32,FA33,FA34,FA35,FA36,CU102,CU65,CU66,CU67,CU68,CU69
         strExc(0) = "SELECT FA04,FA05,FA63,FA64,FA65,FA06,FA07,FA08,FA09,FA52,FA53,FA54" & _
            ",FA29,FA56,FA57,FA58,FA17,FA18,FA19,FA20,FA21,FA22,FA70,FA23,FA32,FA33,FA34,FA35,FA36" & _
            ",FA12,FA14,FA13,FA15 FROM FAGENT WHERE " & ChgFagent(m_strCP44)
         strExc(0) = strExc(0) & " Union SELECT CU04,CU05,CU88,CU89,CU90,CU06,CU58,CU59,CU60,CU61,CU62,CU63" & _
            ",CU79,CU91,CU92,CU93,CU23,CU24,CU25,CU26,CU27,CU28,CU102,CU29,CU65,CU66,CU67,CU68,CU69" & _
            ",CU16,CU18,CU17,CU19 FROM CUSTOMER WHERE " & ChgCustomer(m_strCP44)
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            '若基本檔有聯絡人
            If m_strContact1(0) & m_strContact1(1) & m_strContact1(2) <> "" Then
               If m_strContact1(0) <> "" Then Combo4.AddItem m_strContact1(0)
                If m_strContact1(1) <> "" Then Combo4.AddItem m_strContact1(1)
                If m_strContact1(2) <> "" Then Combo4.AddItem m_strContact1(2)
                Combo4 = m_strContact1(0)
            Else
                'CF聯絡人1
                With RsTemp
                    For i = 6 To 8
                        If IsNull(.Fields(i)) = False And (.Fields(i)) <> "" Then
                            cfa(i) = .Fields(i)
                            Combo4.AddItem cfa(i)
                            Combo4 = cfa(6)
                        End If
                    Next
                End With
            End If
          '若基本檔有聯絡人
          If m_strContact1(0) & m_strContact1(1) & m_strContact1(2) <> "" Then
              If m_strContact2(0) <> "" Then Combo5.AddItem m_strContact2(0)
              If m_strContact2(1) <> "" Then Combo5.AddItem m_strContact2(1)
              If m_strContact2(2) <> "" Then Combo5.AddItem m_strContact2(2)
              Combo5 = m_strContact3(0)
          Else
              'CF聯絡人2
              With RsTemp
                  For i = 9 To 11
                      If IsNull(.Fields(i)) = False And (.Fields(i)) <> "" Then
                          cfa(i) = .Fields(i)
                          Combo5.AddItem cfa(i)
                          Combo5 = cfa(9)
                      End If
                  Next
              End With
          End If
          '若基本檔有實體聯絡人
          If m_strContact3(0) & m_strContact3(1) & m_strContact3(2) <> "" Then
              If m_strContact3(0) <> "" Then Combo6.AddItem m_strContact3(0)
              If m_strContact3(1) <> "" Then Combo6.AddItem m_strContact3(1)
              If m_strContact3(2) <> "" Then Combo6.AddItem m_strContact3(2)
              Combo6 = m_strContact3(0)
          Else
              'CF實體聯絡人
              With RsTemp
                  For i = 13 To 15
                      If IsNull(.Fields(i)) = False And (.Fields(i)) <> "" Then
                          cfa(i) = .Fields(i)
                          Combo6.AddItem cfa(i)
                          Combo6 = cfa(13)
                      End If
                  Next
              End With
          End If
     
          If Combo4.ListCount > 0 And Combo4.ListIndex = -1 Then Combo4.ListIndex = 0
          If Combo5.ListCount > 0 And Combo5.ListIndex = -1 Then Combo5.ListIndex = 0
          If Combo6.ListCount > 0 And Combo6.ListIndex = -1 Then Combo6.ListIndex = 0
     
          With RsTemp
            'CF代理人備註
            If Not IsNull(.Fields(12)) Then cfa(12) = .Fields(12)
            Text6(0) = cfa(12)
            'CF地址
            'Modify by Morgan 2007/1/24 改先抓POB, 加英文地址6
            '中
            cfa(16) = "" & .Fields("FA17")
            '英
            'POB
            If Not IsNull(.Fields("FA32")) Then
               For i = 17 To 21
                  cfa(i) = "" & .Fields(i + 7)
               Next
               cfa(22) = ""
            '地址
            Else
               For i = 17 To 22
                  cfa(i) = "" & .Fields(i)
               Next
            End If
            '日
            cfa(23) = "" & .Fields("FA23")
            'end 2007/1/24
       
            'Add by Morgan 2006/2/13
            m_strFax(1) = "" & .Fields("FA12")
            m_strFax(0) = "" & .Fields("FA14")
            'Add by Morgan 2007/1/19
            m_strFax(3) = "" & .Fields("FA13")
            m_strFax(2) = "" & .Fields("FA15")
          End With
  
         'Modified by Lydia 2020/06/04 備註和各項指示並存(暫時到完全取代), 原本是抓各項指示取代客戶備註
         strMemoY = "": strKeyY = ""
         If strSrvDate(1) >= 各項指示啟用日 Then
             'Added by Lydia 2020/08/26 各項指示：完成確認，各項指示取代原先備註
             If PUB_GetInstConfirm(m_StrUserST03, m_strCP44) = True Then
                 Text6(0) = ""
             End If
             'end 2020/08/26
             If Pub_GetInstructions(Me.Name, m_strCP44, strMemoY, , , , m_strIT10) Then
                 If strMemoY <> "" Then
                     Text6(0) = Text6(0) & IIf(Text6(0) <> "", vbCrLf, "") & strMemoY
                     strKeyY = m_strCP44
                 End If
             End If
         End If
         'end 2020/06/04
      End If
   End If
End Sub

Private Sub Text1_GotFocus()
    TextInverse Text1
    'edit by nickc 2007/06/06 切換輸入法改用API
    'Text1.IMEMode = 2
    CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   'Modify By Sindy 2009/07/24 增加LIN系統類別
   'modify by sonia 2019/7/30 +ACS系統類別
   If Option1(1).Value = True And _
      (Text1.Text = "CFP" Or Text1.Text = "CPS" Or Text1.Text = "FCP" Or _
         Text1.Text = "FG" Or Text1.Text = "CFT" Or Text1.Text = "FCT" Or _
         Text1.Text = "FCL" Or Text1.Text = "CFL" Or Text1.Text = "LIN" Or Text1.Text = "ACS") Then
      txtFaxFace.Enabled = True
   Else
      txtFaxFace.Text = "N"
      txtFaxFace.Enabled = False
   End If
End Sub

'Modify By Sindy 2020/2/11
'Private Sub Text1_Validate(Cancel As Boolean)
Public Sub Text1_Validate(Cancel As Boolean)
'2020/2/11 END
   Dim strTemp1
   Dim strTemp2
   Dim ii As Integer
   Dim jj As Integer
   Dim ss As Integer
   Dim m_Dept As String   '2006/12/11 ADD BY SONIA
    
   '2008/6/25 ADD BY SONIA
   Combo8.Clear
   Combo8.AddItem "一般格式"
   'Modified by Morgan 2015/11/13 改單純下拉
   'Combo8.Text = "一般格式"
   Combo8.ListIndex = 0
   'end 2015/11/13
   '2008/6/25 END

   If Text1.Text = "" Then Exit Sub
   strTemp1 = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
   strTemp2 = Split(Replace(UCase(Text1.Text), ",,", ""), ",")
   For ii = 0 To UBound(strTemp2)
       ss = 0
       For jj = 0 To UBound(strTemp1)
           If strTemp2(ii) = strTemp1(jj) Then
               ss = 1
               Exit For
           End If
       Next jj
       If ss = 0 Then
          '2006/12/11 ADD BY SONIA 開放FF案件之權限
          m_Dept = GetStaffDepartment(strUserNum)
          Select Case m_Dept
            'Modify by Morgan 2007/4/11 加F61
            'Modify by Morgan 2008/4/8 加F81
            'modify by sonia  2014/7/17 加F22
             Case "F21", "F22", "F23", "F61", "F81"  '開放F21,F23使用P,PS,CFP,CPS權限
                If Text1.Text = "P" Or Text1.Text = "PS" Or Text1.Text = "CFP" Or Text1.Text = "CPS" Then
                   Exit For
                End If
             Case "F10", "F11"    '開放F10,F11使用T權限
                'modify by sonia 2014/4/28
                'If Text1.Text = "T" Then
                If InStr(Text1.Text, "T") > 0 Then
                   Exit For
                End If
          End Select
          '2006/12/11 END
          '2010/3/25 ADD BY SONIA 檢查跨部門權限
          If CheckSR09(strUserNum, Text1, "Y", False, Text1, Text2, Text3, Text4) = True Then
             Exit For
          End If
          '2010/3/25 END
          ss = MsgBox(strUserName & " 沒有 " & strTemp2(ii) & " 的權限!! ", , "USER 權限問題")
          Text1.SetFocus
          Text1_GotFocus
          Cancel = True
       End If
   Next ii
End Sub

Private Sub Text2_GotFocus()
    TextInverse Text2
    'edit by nickc 2007/06/06 切換輸入法改用API
    'Text1.IMEMode = 2
    CloseIme
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text3 = "" Then Text3 = "0"
End Sub

Private Sub Text3_GotFocus()
    TextInverse Text3
    'edit by nickc 2007/06/06 切換輸入法改用API
    'Text1.IMEMode = 2
    CloseIme
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If Text3 = "" Then Text3 = "0"
   If Text4 = "" Then Text4 = "00"
End Sub

Private Sub Text4_GotFocus()
    TextInverse Text4
    'edit by nickc 2007/06/06 切換輸入法改用API
    'Text1.IMEMode = 2
    CloseIme
End Sub

Private Sub Text4_LostFocus()
   Dim strMsg As String 'Add by Amy 2016/04/12
   
   If Text3 = "" Then Text3 = "0"
   If Text4 = "" Then Text4 = "00"
   '2014/7/17 ADD BY SONIA 因開放外專程序可抓FMP資料,故再檢查跨部門案件權限
   If CheckSR09(strUserNum, Text1, "Y", True, Text1, Text2, Text3, Text4) = False Then
      Exit Sub
   End If
   '2014/7/17 END
   Read
   'Add by Amy 2016/04/12 +CFT案且操作人員為F1開頭部門人員,若案件地址與客戶地址不同時彈訊息
   If Text1 = "CFT" And Left(GetStaffDepartment(strUserNum), 2) = "F1" Then
        If ChkCaseAndCusAddrNotAlike(strMsg) = True Then
            MsgBox strMsg, vbCritical, "檢核資料"
        End If
   End If
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
   If Text3 = "" Then Text3 = "0"
   If Text4 = "" Then Text4 = "00"
End Sub

'Modify By Sindy 2014/10/3
'Private Sub Read()
Public Sub Read()
'2014/10/3 END
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim bolFMPCase As Boolean 'Add By Sindy 2014/8/21
   
   m_bolReadOK = False 'Added by Morgan 2013/3/8
   m_iLang = -1
   txtConDepJP = ""
   
   Option2.Value = False
   Option3.Value = False
   Option4.Value = False
   Option5.Value = False
   Option6.Value = False
   Combo1.Clear
   Combo2.Clear
   Combo3.Clear
   Combo4.Clear
   Combo5.Clear
   Combo6.Clear
   Combo7.Clear
   Text5 = ""
   Text6(0) = ""
   Text6(1) = ""
   'Add By Cheng 2002/04/29
   lblClose.Caption = ""
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   m_AboutDeadLine = "" 'Added by Morgan 2015/5/14
   
   'Add By Sindy 2014/8/13
   bolFMPCase = False
'   bolIsECase = False 'Add By Sindy 2014/8/26
   Command1.Default = True
   lblSendMailDt.Visible = False 'Add By Sindy 2018/5/11
   cmdFCMail(0).Enabled = False: cmdFCMail(1).Enabled = False
   cmdFCMail(2).Enabled = False 'Added by Lydia 2015/10/30
   strExc(0) = "select cp12 from CaseProgress" & _
               " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
               " and cp09=(select max(cp09) from CaseProgress" & _
               " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
               " and cp05=(select max(cp05) from CaseProgress" & _
               " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "')" & _
               ")"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      '1.該案件的最大收文日最大收文號的業務區為F2字頭的
      If Left(RsTemp.Fields("cp12"), 2) = "F2" Then
         If pa(1) = "FG" Or pa(1) = "CFP" Or pa(1) = "CPS" Or pa(1) = "P" Or pa(1) = "PS" Then
            bolFMPCase = True
         End If
'         strExc(0) = "select pa75,pa26,pa142,fa86,cu124 from patent,fagent,customer" & _
'                     " where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "'" & _
'                     " and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+)" & _
'                     " and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+)" & _
'              " union select sp26,sp08,sp80,fa86,cu124 from servicepractice,fagent,customer" & _
'                     " where sp01='" & pa(1) & "' and sp02='" & pa(2) & "' and sp03='" & pa(3) & "' and sp04='" & pa(4) & "'" & _
'                     " and substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+)" & _
'                     " and substr(sp26,1,8)=fa01(+) and substr(sp26,9,1)=fa02(+)"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         '若PA142或FA86或CU124有一個設定為Y以E-Mail通知者,就屬於E化的案件
'         If intI = 1 Then
'            If "" & RsTemp.Fields("pa142") = "Y" Or _
'               "" & RsTemp.Fields("fa86") = "Y" Or _
'               "" & RsTemp.Fields("cu124") = "Y" Then
'               bolIsECase = True 'Add By Sindy 2014/8/26 為E化案件
''               '操作人員的部門為F22外專程序或M51電腦中心人員才能使用
''               If m_StrUserSt03 = "M51" Or m_StrUserSt03 = "F22" Then
''                  cmdFCMail.Enabled = True
''                  cmdFCMail.Default = True
''               End If
'            End If
'         End If
         
         'Modify By Sindy 2014/8/26 非E化案件也可以使用outlook
         '操作人員的部門為F22外專程序或M51電腦中心人員才能使用
         'Modify By Sindy 2014/9/15 開放F23外專承辦也可以使用
         If m_StrUserST03 = "M51" Or m_StrUserST03 = "F22" Or m_StrUserST03 = "F23" Then
            cmdFCMail(0).Enabled = True: cmdFCMail(1).Enabled = True
            cmdFCMail(0).Default = True
         End If
         'Added by Lydia 2015/10/30
         'Modified by Lydia 2016/07/04
         'If Pub_StrUserSt03 = "M51" Or InStr("73023,82045", strUserNum) > 0 Then
         'Mark by Lydia 2025/03/13 已不再使用
         'If cmdFCMail(2).Visible = True Then
         '   cmdFCMail(2).Enabled = True
         'End If
         'end 2025/03/13
      End If
   End If
   '2014/8/13 END
   
   m_strCP44 = "": m_strFCAgent = "": m_CustNo(1) = ""
   Erase m_strContact1: Erase m_strContact2: Erase m_strContact3
   'Add by Morgan 2005/6/6 清除全域變數
   Erase cfa: Erase m_strFax: Erase fa: Erase cu
   
   AppNo = "": casetype4 = "": casetype5 = ""  '2008/6/25 add by sonia
   m_Zip = "" 'Add by Morgan 2008/7/17
   m_ChildCP44 = "" 'Added by Lydia 2019/02/26
   
   m_QCase = Trim(Text1) & Trim(Text2) & Trim(Text3) & Trim(Text4) 'Add By Sindy 2013/2/18 記錄目前所查詢的案號資料
   Select Case pa(1) '判斷系統類別
      Case "P", "CFP", "FCP" '專利
          'edit by nickc 2007/02/02 不用 dll 了
          'If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
          If ClsPDReadPatentDatabase(pa(), intWhere) Then
            m_bolReadOK = True 'Added by Morgan 2013/3/8
              '若有基本資料
              If Not IsNull(pa()) Then
                  'Added by Lydia 2019/11/01 利益衝突案件：逐案號判斷
                  If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
                      If PUB_ChkCufaByCase(Me.Name, pa(1), pa(1) & pa(2) & pa(3) & pa(4), "" & pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30), "" & pa(75)) = False Then
                          MsgBox MsgText(1109), vbInformation, MsgText(1110)
                          GoTo JumpToExitPA
                      End If
                  End If
                  'end 2019/11/01
                  
                  '案件名稱
                  If IsNull(pa(5)) = False And pa(5) <> "" Then Combo1.AddItem pa(5)
                  If IsNull(pa(6)) = False And pa(6) <> "" Then Combo1.AddItem pa(6)
                  If IsNull(pa(7)) = False And pa(7) <> "" Then Combo1.AddItem pa(7)
                  '2008/6/25 add by sonia
                  '2008/12/1 MODIFY BY SONIA 有申請案號才帶,無則不帶
                  'If IsNull(pa(11)) = False And pa(11) <> "" Then AppNo = pa(11)
                  If IsNull(pa(11)) = False And pa(11) <> "" Then AppNo = "第" & pa(11) & "號"
                  If IsNull(pa(48)) = False And pa(48) <> "" Then casetype4 = "(" & pa(48) & ")"
                  '2008/6/25 end
                  '2008/9/18 ADD BY SONIA
                  If IsNull(pa(47)) = False And pa(47) <> "" Then casetype5 = pa(47)
                  '2008/9/18 END
                  'Add By Cheng 2002/04/29
                  '是否閉卷
                  If Len("" & pa(57)) <= 0 Then
                     lblClose.Caption = ""
                  Else
                     lblClose.Caption = "已閉卷"
                  End If
                  '抓國家名稱
                  Label11 = pa(9)
                  'edit by nickc 2007/02/02 不用 dll 了
                  'If objPublicData.GetNation(pA(9), strTemp) Then Label12.Caption = strTemp
                  If ClsPDGetNation(pa(9), strTemp) Then Label12.Caption = strTemp
                  Combo1 = pa(5)
             
                 'Add by Morgan 2007/3/12
                 '基本檔定稿語文
                 If pa(85) <> "" Then
                    m_iLang = Val(pa(85)) - 1
                 End If
                 'end 2007/3/12
            
                  '抓FC代理人資料
                  strExc(0) = "SELECT FA04,FA05,FA63,FA64,FA65,FA06,FA31 FROM FAGENT WHERE " & ChgFagent(pa(75))
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                      With RsTemp
                          For i = 0 To 5
                              If IsNull(.Fields(i)) = False And (.Fields(i)) <> "" Then
                                  fa(i) = .Fields(i)
                                  Combo2.AddItem fa(i)
                                  Combo2 = fa(0)
                              End If
                          Next
                          'Add by Morgan 2007/3/12
                          '基本檔定稿語文
                          If m_iLang = -1 And Not IsNull(.Fields("FA31")) Then
                             m_iLang = Val(.Fields("FA31")) - 1
                          End If
                          'end 2007/3/12
                      End With
                  End If
                  
                  'Added by Lydia 2019/02/26 CFP之EPC子案於撰寫信函抓指定代理人(與期限通知管制表及年費發文指示信之子案代理人抓法相同)
                  If pa(1) = "CFP" And pa(4) <> "00" Then
                      strExc(0) = "SELECT PA01,PA02,PA03,PA04,PA09 FROM PATENT WHERE PA01='" & pa(1) & "' AND PA02='" & pa(2) & "' AND PA03='0' AND PA04='00' "
                      intI = 1
                      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                      If intI = 1 Then
                          If "" & RsTemp.Fields("PA09") = "221" Then
                               If PUB_GetEPCtoCP44(pa(1), pa(2), pa(3), pa(4), m_ChildCP44, strExc(1)) = True Then
                               End If
                          End If
                      End If
                  End If
                  'end 2019/02/26
                  '抓CF代理人資料
                  If bolFA = True Then
                      '2006/5/23 MODIFY BY SONIA 加 CP09<'C'條件
                      'Modify by Morgan 2008/2/18 改以發文日排序
                      '2009/6/23 MODIFY BY SONIA 郭雅娟說取消發文日CP27>0限制,因大陸已輸指示信未發文前無法使用
                      'Modify by Morgan 2011/1/5 +CP57 is null
                      'Modified by Morgan 2023/10/30
                      'StrSQLa = "Select CP44 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                                      " AND CP44 Is Not Null AND CP09<'C' and CP57 is null Order By CP27 Desc, CP09 Desc "
                      'rsA.CursorLocation = adUseClient
                      'rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                      'If rsA.RecordCount > 0 Then
                      '  m_strCP44 = "" & rsA.Fields(0).Value
                      'Else
                      '    m_strCP44 = ""
                      'End If
                      If m_IDSCP09 <> "" Then
                        strExc(0) = "select cp44 from caseprogress where cp09='" & m_IDSCP09 & "' and cp44 is not null"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           m_strCP44 = RsTemp(0)
                        Else
                           m_strCP44 = ""
                        End If
                      'Modified by Morgan 2024/3/27 +m_strCP45,m_strCP09
                      ElseIf ClsPDGetCasePreAgent(pa(), m_strCP44, False, m_strCP45, m_strCP09) = False Then
                        If ClsPDGetCasePreAgent(pa(), m_strCP44, False, m_strCP45, m_strCP09, False) = False Then 'Added by Morgan 2025/4/25 沒有代理人時就不排除年費 Ex:CFP-034575
                           m_strCP44 = ""
                        End If
                      End If
                      'end 2023/10/30
                      
                      'Added by Morgan 2025/2/26
                      '若有聯絡人需排除否則會抓不到代理人名稱
                      intI = InStr(m_strCP44, "-")
                      If intI > 0 Then
                        m_strCP44 = Left(m_strCP44, intI - 1)
                      End If
                      'end 2025/2/26
                      
                      If m_ChildCP44 <> "" Then m_strCP44 = m_ChildCP44 'Added by Lydia 2019/02/26
                      If rsA.State <> adStateClosed Then rsA.Close
                      Set rsA = Nothing
                      SetCFRef 'Modified by Morgan 2023/12/13 原程式改為函數以便共用
                  End If
                  
                 'Add by Morgan 2007/3/12 FCP預設FC代理人及語文
                 'Modify By Sindy 2014/8/21 FMP的案件設定值同FCP
                 'If pa(1) = "FCP" Then
                 If pa(1) = "FCP" Or bolFMPCase = True Then
                 '2014/8/21 END
                    If m_iLang = -1 Then m_iLang = 1
                    If Option1(m_iLang).Value = True Then
                       Option1_Click m_iLang
                    Else
                       Option1(m_iLang).Value = True
                    End If
          
                    If pa(75) <> "" Then
                       If Option2.Value = True Then
                          Option2_Click
                       Else
                          Option2.Value = True
                       End If
                    End If
                    Option4.Value = True
                    Select Case m_iLang
                       Case 0 '中->英->日
                          '案件名稱
                          If pa(5) <> "" Then
                             Combo1.Text = pa(5)
                          ElseIf pa(6) <> "" Then
                             Combo1.Text = pa(6)
                          Else
                             Combo1.Text = pa(7)
                          End If
                          '聯絡人
                          If m_strContact1(0) & m_strContact1(1) & m_strContact1(2) & m_strContact2(0) & m_strContact2(1) & m_strContact2(2) <> "" Then
                             '1
                             If m_strContact1(0) <> "" Then
                                Combo4.Text = m_strContact1(0)
                             ElseIf m_strContact1(1) <> "" Then
                                Combo4.Text = m_strContact1(1)
                             Else
                                Combo4.Text = m_strContact1(2)
                             End If
                             '2
                             If m_strContact2(0) <> "" Then
                                Combo5.Text = m_strContact2(0)
                             ElseIf m_strContact2(1) <> "" Then
                                Combo5.Text = m_strContact2(1)
                             Else
                                Combo5.Text = m_strContact2(2)
                             End If
                          Else
                             '1
                             If fa(6) <> "" Then
                                Combo4.Text = fa(6)
                             ElseIf fa(7) <> "" Then
                                Combo4.Text = fa(7)
                             Else
                                Combo4.Text = fa(8)
                             End If
                             '2
                             If fa(9) <> "" Then
                                Combo5.Text = fa(9)
                             ElseIf fa(10) <> "" Then
                                Combo5.Text = fa(10)
                             Else
                                Combo5.Text = fa(11)
                             End If
                          End If
                
                       Case 1 '英->中->日
                          '案件名稱
                          If pa(6) <> "" Then
                             Combo1.Text = pa(6)
                          ElseIf pa(5) <> "" Then
                             Combo1.Text = pa(5)
                          Else
                             Combo1.Text = pa(7)
                          End If
                          '聯絡人
                          If m_strContact1(0) & m_strContact1(1) & m_strContact1(2) & m_strContact2(0) & m_strContact2(1) & m_strContact2(2) <> "" Then
                             '1
                             If m_strContact1(1) <> "" Then
                                Combo4.Text = m_strContact1(1)
                             ElseIf m_strContact1(0) <> "" Then
                                Combo4.Text = m_strContact1(0)
                             Else
                                Combo4.Text = m_strContact1(2)
                             End If
                             '2
                             If m_strContact2(1) <> "" Then
                                Combo5.Text = m_strContact2(1)
                             ElseIf m_strContact2(0) <> "" Then
                                Combo5.Text = m_strContact2(0)
                             Else
                                Combo5.Text = m_strContact2(2)
                             End If
                          Else
                             '1
                             If fa(7) <> "" Then
                                Combo4.Text = fa(7)
                             ElseIf fa(6) <> "" Then
                                Combo4.Text = fa(6)
                             Else
                                Combo4.Text = fa(8)
                             End If
                             '2
                             If fa(10) <> "" Then
                                Combo5.Text = fa(10)
                             ElseIf fa(9) <> "" Then
                                Combo5.Text = fa(9)
                             Else
                                Combo5.Text = fa(11)
                             End If
                          End If
                
                       Case 2 '日->英->中
                          '案件名稱
                          If pa(7) <> "" Then
                             Combo1.Text = pa(7)
                          ElseIf pa(6) <> "" Then
                             Combo1.Text = pa(6)
                          Else
                             Combo1.Text = pa(5)
                          End If
                          '聯絡人
                          If m_strContact1(0) & m_strContact1(1) & m_strContact1(2) & m_strContact2(0) & m_strContact2(1) & m_strContact2(2) <> "" Then
                             '1
                             If m_strContact1(2) <> "" Then
                                Combo4.Text = m_strContact1(2)
                             ElseIf m_strContact1(1) <> "" Then
                                Combo4.Text = m_strContact1(1)
                             Else
                                Combo4.Text = m_strContact1(0)
                             End If
                             '2
                             If m_strContact2(2) <> "" Then
                                Combo5.Text = m_strContact2(2)
                             ElseIf m_strContact2(1) <> "" Then
                                Combo5.Text = m_strContact2(1)
                             Else
                                Combo5.Text = m_strContact2(0)
                             End If
                          Else
                             '1
                             If fa(8) <> "" Then
                                Combo4.Text = fa(8)
                             ElseIf fa(7) <> "" Then
                                Combo4.Text = fa(7)
                             Else
                                Combo4.Text = fa(6)
                             End If
                             '2
                             If fa(11) <> "" Then
                                Combo5.Text = fa(11)
                             ElseIf fa(10) <> "" Then
                                Combo5.Text = fa(10)
                             Else
                                Combo5.Text = fa(9)
                             End If
                          End If
                    End Select
                    'FC代理人
                    If Option2.Value = True Then
                       Select Case m_iLang
                          Case 0 '中->英->日
                             If fa(0) <> "" Then
                                Combo2.Text = fa(0)
                             ElseIf fa(1) <> "" Then
                                Combo2.Text = fa(1)
                             Else
                                Combo2.Text = fa(5)
                             End If
                   
                          Case 1 '英->中->日
                             If fa(1) <> "" Then
                                Combo2.Text = fa(1)
                             ElseIf fa(0) <> "" Then
                                Combo2.Text = fa(0)
                             Else
                                Combo2.Text = fa(5)
                             End If
                   
                          Case 2 '日->英->中
                             If fa(5) <> "" Then
                                Combo2.Text = fa(5)
                             ElseIf fa(1) <> "" Then
                                Combo2.Text = fa(1)
                             Else
                                Combo2.Text = fa(0)
                             End If
                       End Select
                    End If
                 End If
                 'end 2007/3/12
            
                  '抓客戶基本檔
                  'Modify by Morgan 2011/9/16 +客戶備註
                  strExc(0) = "SELECT CU04,CU05,CU88,CU89,CU90,CU06,CU79 FROM CUSTOMER WHERE " & ChgCustomer(pa(26))
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                      With RsTemp
                          For i = 0 To 5
                              If IsNull(.Fields(i)) = False And (.Fields(i)) <> "" Then
                                  cu(i) = .Fields(i)
                                  Combo3.AddItem cu(i)
                                  Combo3 = cu(0)
                              End If
                          Next
                          Text6(1) = "" & .Fields("CU79")
                      End With
                  End If
                
                  'Modified by Lydia 2020/06/04 備註和各項指示並存(暫時到完全取代), 原本是抓各項指示取代客戶備註
                  strMemoX = "": strKeyX = ""
                  If strSrvDate(1) >= 各項指示啟用日 Then
                      'Added by Lydia 2020/08/26 各項指示：完成確認，各項指示取代原先備註
                      If PUB_GetInstConfirm(m_StrUserST03, pa(26)) = True Then
                           Text6(1) = ""
                      End If
                      'end 2020/08/26
                      If Pub_GetInstructions(Me.Name, pa(26), strMemoX, , , , m_strIT10) Then
                          If strMemoX <> "" Then
                             Text6(1) = Text6(1) & IIf(Text6(1) <> "", vbCrLf, "") & strMemoX
                             strKeyX = pa(26)
                          End If
                      End If
                 End If
                 'end 2020/06/04
                 
                  '案件備註
                  Text5 = pa(91)
              End If
               'add by sonia 2019/11/19
               bolDateType = DateType(pa(9), pa(26), pa(75))
               'end 2019/11/19
          Else
JumpToExitPA: 'Added by Lydia 2019/11/01
              Text1.SetFocus
              Text1_GotFocus
               'Added by Morgan 2013/3/8
              '案號錯誤時應跳離否則否可能會重複觸發事件導致無窮回圈
               Exit Sub
          End If
         
         '2008/6/25 ADD BY SONIA
         Combo8.Clear
         Combo8.AddItem "一般格式"
         'Modified by Morgan 2015/11/13 改單純下拉
         'Combo8.Text = "一般格式"
         Combo8.ListIndex = 0
         'end 2015/11/13
         
         Select Case Text1.Text
            'Added by Morgan 2024/3/29
            Case "FCP"
               If OutCallCP09 <> "" Then
                  StrSQLa = "select cp09,cp10,cpm03 from caseprogress,casepropertymap where cp09='" & OutCallCP09 & "' and cpm01(+)=cp01 and cpm02(+)=cp10"
                  rsA.CursorLocation = adUseClient
                  rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                  If rsA.RecordCount > 0 Then
                     If pa(150) <> "3" And rsA("cp10") = "926" Then
                        Combo8.AddItem rsA("cpm03") & "     " & rsA("cp10")
                        Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0))
                        Combo8.ListIndex = Combo8.ListCount - 1
                     End If
                  End If
               Else
                  StrSQLa = "select cp09,cp10,cpm03 from caseprogress,casepropertymap where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " and cp10='926' and cpm01(+)=cp01 and cpm02(+)=cp10 order by nvl(cp27,cp05) desc"
                  rsA.CursorLocation = adUseClient
                  rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                  If rsA.RecordCount > 0 Then
                     Combo8.AddItem rsA("cpm03") & "     " & rsA("cp10")
                     Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0))
                  End If
               End If
            Case "CFP"
               '核駁
               StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                               " AND CP10='1002' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  Combo8.AddItem "核駁　　　　　　1002"
                  Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               '最終核駁
               StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                               " AND CP10='1006' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  Combo8.AddItem "最終核駁　　　　1006"
                  Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               '通知要求選取
               StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                               " AND CP10='1206' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  Combo8.AddItem "通知要求選取　　1206"
                  Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               '檢索報告
               StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                               " AND CP10='1209' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  Combo8.AddItem "檢索報告　　　　1209"
                  Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               
               'Add by Morgan 2010/7/19
               '國際初步審查報告
               StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                               " AND CP10='1216' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  Combo8.AddItem "國際初步審查報告　　　　1216"
                  Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               'end 2010/7/19
               
               '2013/1/21 ADD BY SONIA
               '分析(同核駁格式)
               StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                               " AND CP10='941' AND CP27 IS NULL AND CP57 IS NULL AND CP09<'C' "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  Combo8.AddItem "分析　　　　　　941"
                  Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               
            '2008/12/1 ADD BY SONIA
            Case "P"
               '告知代理人        函知客戶檢索報告用
               StrSQLa = "Select C1.CP09 FROM CaseProgress C1,CaseProgress C2 WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                               " AND C1.CP10='901' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND SUBSTR(C1.CP09,1,1)='B' AND C1.CP43=C2.CP09(+) AND '421'=C2.CP10(+) AND C2.CP47 IS NOT NULL"
               'Add by Morgan 2009/10/2 +大陸檢索報告改由核准輸入
               '2009/12/3 MODIFY BY SONIA 取消CP47限制
               'StrSQLa = StrSQLa & " union Select C1.CP09 FROM CaseProgress C1,CaseProgress C2 WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                               " AND C1.CP10='1001' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP43=C2.CP09(+) AND C2.CP10 in ('421','426') AND C2.CP47 IS NOT NULL"
               'Modified by Morgan 2012/4/20 +423
               'MODIFY BY SONIA 2014/6/4 423專利權評價報告改獨立出來加在下面P-105534
                'Modified by Lydia 2015/10/05 + 1008
               StrSQLa = StrSQLa & " union Select C1.CP09 FROM CaseProgress C1,CaseProgress C2 WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                               " AND C1.CP10 in ('1001','1008') AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP43=C2.CP09(+) AND C2.CP10 in ('421','426')"
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  Combo8.AddItem "檢索報告　　　　1209"
                  Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               'ADD BY SONIA 2014/6/4 423專利權評價報告改獨立出來P-105534
                'Modified by Lydia 2015/10/05 + 1008
               StrSQLa = StrSQLa & " union Select C1.CP09 FROM CaseProgress C1,CaseProgress C2 WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                               " AND C1.CP10 in ('1001','1008') AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP43=C2.CP09(+) AND C2.CP10='423'"
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  Combo8.AddItem "專利權評價報告  1209"
                  Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               'END 2014/6/4
               '2009/3/3 ADD BY SONIA 加PCT案之檢索報告1209及國際初步審查報告1216
               'Modified by Morgan 2015/2/26 + AND C1.CP10 IN ('1209','1216')
               StrSQLa = "Select C2.CP10,C2.CP09 FROM CaseProgress C1,CaseProgress C2 WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                               " AND C1.CP10='901' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND SUBSTR(C1.CP09,1,1)='B' AND C1.CP43=C2.CP09(+)  AND C2.CP10 IN ('1209','1216')"
               'Added by Morgan 2015/2/26
               'PCT檢索報告1209,國際初步審查報告1216改不自動發文及新增告代(確定無告代未發文後可省略原語法)
               StrSQLa = StrSQLa & " union Select CP10,CP09 FROM CaseProgress C1 WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                               " AND C1.CP10 IN ('1209','1216') AND C1.CP27 IS NULL AND C1.CP57 IS NULL"
               'end 2015/2/26
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  If rsA.Fields(0) = "1209" Then
                     Combo8.AddItem "檢索報告　　　　1209"
                  ElseIf rsA.Fields(0) = "1216" Then
                     Combo8.AddItem "國際初步審查報告1216"
                  End If
                  Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(1)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               '2009/3/3 END
               
            '2008/12/1 END
               
               'Add By Sindy 2013/3/29
               '台灣舉發及答辯審定書
               If Trim(Label11) = "000" Then
                  intI = 0 'Added by Morgan 2015/2/10
                  'Modified by Morgan 2022/8/24 +部分准駁1009
                  StrSQLa = "Select C1.CP43,C1.CP10 FROM CaseProgress C1,CaseProgress C2,CaseProgress C3 WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                            " AND C1.CP10='941' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP43=C2.CP09(+) AND C2.CP10 in ('1001','1002','1009','1503') AND C2.CP43=C3.CP09(+) AND C3.CP10 in ('803','804')"
                            
                  'Added by Morgan 2016/3/16
                  'P臺灣案的答辯或舉發答辯的審定來函改工程師承辦(原內部收文分析)
                  'Modified by Morgan 2022/8/24 +部分准駁1009
                  StrSQLa = StrSQLa & " union Select CP09,CP10 FROM CaseProgress a WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' and cp10 IN ('1001','1002','1009','1503') and cp27||cp57 is null" & _
                        " and exists(select * from CaseProgress b where b.CP09=a.CP43 AND b.CP10 in ('803','804'))"
                  'end 2016/3/16
                  intI = 1
                  Set rsA = ClsLawReadRstMsg(intI, StrSQLa)
                  If intI = 1 Then
                     'Modified by Morgan 2016/3/18
                     'Combo8.AddItem "分析　　　　　　941"
                     Combo8.AddItem "分析　　　　　　" & rsA("cp10")
                     'end 2016/3/10
                     Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
                  End If
                  If rsA.State <> adStateClosed Then rsA.Close
                  
                  'Added by Morgan 2015/2/10
                  If intI = 0 Then
                     'Modified by Morgan 2015/5/14 +若相關收文號有期限時信函要帶出
                     'Modified by Morgan 2021/10/19 +未發文的OA也同分析出通用定稿(配合顧服組特定客戶的案件要用OA跑歷程)
                     'Modified by Morgan 2022/4/11 NP還要串本所案號否則會抓到IDS期限 Ex:P127778
                     StrSQLa = "Select CP09,CPM03,NP08,to_char(to_date(np08,'yyyymmdd')-14,'yyyymmdd') dt,CP10 FROM CaseProgress C1,nextprogress,casepropertymap WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                            " AND C1.CP10='941' AND C1.CP27 IS NULL AND C1.CP57 IS NULL and np01(+)=cp43 and np02(+)=cp01 and np03(+)=cp02 and np04(+)=cp03 and np05(+)=cp04 and np06(+) is null and cpm01(+)=np02 and cpm02(+)=np07" & _
                            " union Select CP09,CPM03,NP08,to_char(to_date(np08,'yyyymmdd')-14,'yyyymmdd') dt,CP10 FROM CaseProgress C1,nextprogress,casepropertymap WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                            " AND C1.CP10 in (" & PatentOAPtyList & ") AND C1.CP27 IS NULL AND C1.CP57 IS NULL and np01(+)=cp09 and np02(+)=cp01 and np03(+)=cp02 and np04(+)=cp03 and np05(+)=cp04 and np06(+) is null and cpm01(+)=np02 and cpm02(+)=np07"
                     rsA.CursorLocation = adUseClient
                     rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                     If rsA.RecordCount > 0 Then
                        'Added by Morgan 2015/5/14
                        If rsA("np08") > 0 Then
                           'modify by sonia 2020/2/20 TranslateKeyWord(incCNV_CHINESE_MINKO...加傳本所案號,以判斷日期欄之民國或西元格式
                           m_AboutDeadLine = "　　本案應於" & TranslateKeyWord(incCNV_CHINESE_MINKO, rsA("np08"), Empty, pa(1) & pa(2) & pa(3) & pa(4)) & "以前提出" & rsA("cpm03") & "，若欲辦理，請提早於" & TranslateKeyWord(incCNV_CHINESE_MINKO, rsA("dt"), Empty, pa(1) & pa(2) & pa(3) & pa(4)) & "前通知本所，如有任何問題，請隨時不吝賜教，本所將竭盡全力，為　貴單位服務。"
                        End If
                        'end 2015/5/14
                        
                        'Added by Morgan 2021/10/19
                        If Left(rsA("cp09"), 1) = "C" Then
                           ClsPDGetCaseProperty "P", rsA("cp10"), strExc(1)
                           If Len(strExc(1)) < 8 Then
                              strExc(1) = strExc(1) & String(8 - Len(strExc(1)), "　")
                           End If
                           Combo8.AddItem strExc(1) & rsA("cp10") & " " '注意後面有多一個空白以區別為用一般通用定稿
                        Else
                        'end 2021/10/19
                           Combo8.AddItem "分析　　　　　　941 " '注意後面有多一個空白以區別為用一般通用定稿
                        End If 'Added by Morgan 2021/10/19
                        
                        Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
                     End If
                     If rsA.State <> adStateClosed Then rsA.Close
                  End If
                  'end 2015/2/10
                  
                  Set rsA = Nothing
               End If
               '2013/3/29 End
         End Select
            '2008/6/25 END
       
      Case "T", "TF", "CFT", "FCT" '商標
          'edit by nickc 2007/02/02 不用 dll 了
          'If objPublicData.ReadTrademarkDatabase(pA(), intWhere) Then
          If ClsPDReadTrademarkDatabase(pa(), intWhere) Then
            m_bolReadOK = True 'Added by Morgan 2013/3/8
              '若有案件基本資料
              If Not IsNull(pa()) Then
                  'Added by Lydia 2019/11/01 利益衝突案件：逐案號判斷
                  If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
                      If PUB_ChkCufaByCase(Me.Name, pa(1), pa(1) & pa(2) & pa(3) & pa(4), "" & pa(23) & "," & pa(78) & "," & pa(79) & "," & pa(80) & "," & pa(81), "" & pa(44)) = False Then
                          MsgBox MsgText(1109), vbInformation, MsgText(1110)
                          GoTo JumpToExitTM
                      End If
                  End If
                  'end 2019/11/01
                  
                  '案件名稱
                  If IsNull(pa(5)) = False And pa(5) <> "" Then Combo1.AddItem pa(5)
                  If IsNull(pa(6)) = False And pa(6) <> "" Then Combo1.AddItem pa(6)
                  If IsNull(pa(7)) = False And pa(7) <> "" Then Combo1.AddItem pa(7)
                  '2008/6/25 add by sonia
                  If IsNull(pa(35)) = False And pa(35) <> "" Then casetype4 = "(" & pa(35) & ")"
                  '2008/6/25 end
                  '2008/9/18 ADD BY SONIA "第" & pa(11) & "號"
                  '2008/12/1 MODIFY BY SONIA
                  'If IsNull(pa(12)) = False And pa(12) <> "" Then AppNo = pa(12)
                  If IsNull(pa(12)) = False And pa(12) <> "" Then AppNo = "第" & pa(12) & "號"
                  '若有審定號則改抓審定號
                  '2008/12/1 MODIFY BY SONIA
                  'If IsNull(pa(15)) = False And pa(15) <> "" Then AppNo = pa(15)
                  If IsNull(pa(15)) = False And pa(15) <> "" Then AppNo = "第" & pa(15) & "號"
                  If IsNull(pa(34)) = False And pa(34) <> "" Then casetype5 = pa(34)
                  '2008/9/18 END
                  'Add By Cheng 2002/04/29
                  '是否閉卷
                  If Len("" & pa(29)) <= 0 Then
                      lblClose.Caption = ""
                  Else
                      lblClose.Caption = "已閉卷"
                  End If
                  '申請國家
                  Label11 = pa(10)
                  'edit by nickc 2007/02/02 不用 dll 了
                  'If objPublicData.GetNation(pA(10), strTemp) Then
                  If ClsPDGetNation(pa(10), strTemp) Then
                      Label12.Caption = strTemp
                  End If
                  Combo1 = pa(5)
                  '抓客戶基本檔
                  'Modify by Morgan 2011/9/16 +CU79
                  strExc(0) = "SELECT CU04,CU05,CU88,CU89,CU90,CU06,CU79 FROM CUSTOMER WHERE " & ChgCustomer(pa(23))
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                      With RsTemp
                          For i = 0 To 5
                              If IsNull(.Fields(i)) = False And (.Fields(i)) <> "" Then
                                  cu(i) = .Fields(i)
                                  Combo3.AddItem cu(i)
                                  Combo3 = cu(0)
                              End If
                          Next
                          Text6(1) = "" & .Fields("CU79")
                      End With
                  End If

                   'Modified by Lydia 2020/06/04 備註和各項指示並存(暫時到完全取代), 原本是抓各項指示取代客戶備註
                   strMemoX = "": strKeyX = ""
                   If strSrvDate(1) >= 各項指示啟用日 Then
                      'Added by Lydia 2020/08/26 各項指示：完成確認，各項指示取代原先備註
                      If PUB_GetInstConfirm(m_StrUserST03, pa(23)) = True Then
                           Text6(1) = ""
                      End If
                      'end 2020/08/26
                      If Pub_GetInstructions(Me.Name, pa(23), strMemoX, , , , m_strIT10) Then
                          If strMemoX <> "" Then
                             Text6(1) = Text6(1) & IIf(Text6(1) <> "", vbCrLf, "") & strMemoX
                             strKeyX = pa(23)
                          End If
                      End If
                   End If
                   'end 2020/06/04
                 
                  '抓FCAgent
                  strExc(0) = "SELECT FA04,FA05,FA63,FA64,FA65,FA06 FROM FAGENT WHERE " & ChgFagent(pa(44))
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                      With RsTemp
                          For i = 0 To 5
                              If IsNull(.Fields(i)) = False And (.Fields(i)) <> "" Then
                                  fa(i) = .Fields(i)
                                  Combo2.AddItem fa(i)
                                  Combo2 = fa(0)
                              End If
                          Next
                      End With
                  End If
                  'Add By Cheng 2003/06/26
                  '抓CFAgent
                  'Modify by Morgan 2004/10/7
                  'If bolFNation = True Then
                  m_Y52269FA16Mail = "" 'CF代理人代表信箱 Add By Sindy 2017/3/7
                  If bolFA = True Then
                      '2006/5/23 MODIFY BY SONIA 加 CP09<'C'條件
                      '2009/6/23 MODIFY BY SONIA 郭雅娟說取消發文日CP27>0限制,因大陸已輸指示信未發文前無法使用
                      'Modify by Morgan 2011/1/5 +CP57 is null
                      StrSQLa = "Select CP44 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                                      " AND CP44 Is Not Null AND CP09<'C' and cp57 is null"
                      'Modify By Sindy 2013/3/25 CFT撰寫信函之CF代理人名稱：文件簽證711及申請英文證明304 不要列入
                      If pa(1) = "CFT" Then
                        StrSQLa = StrSQLa & " AND CP10 NOT IN ('711','304')"
                      End If
                      '2013/3/25 End
                      StrSQLa = StrSQLa & " Order By CP27 Desc, CP09 Desc"
                      rsA.CursorLocation = adUseClient
                      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                      If rsA.RecordCount > 0 Then
                          m_strCP44 = "" & rsA.Fields(0).Value
                      Else
                          m_strCP44 = ""
                      End If
                      If rsA.State <> adStateClosed Then rsA.Close
                      Set rsA = Nothing
                      'Modify By Sindy 2017/3/7 + FA16
                      strExc(0) = "SELECT FA04,FA05,FA63,FA64,FA65,FA06,FA16 FROM FAGENT WHERE " & ChgFagent(m_strCP44)
                      '抓CFagent
                      intI = 1
                      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                      If intI = 1 Then
                          With RsTemp
                              m_Y52269FA16Mail = "" & .Fields("FA16") 'CF代理人代表信箱 Add By Sindy 2017/3/7
                              For i = 0 To 5
                                  If IsNull(.Fields(i)) = False And (.Fields(i)) <> "" Then
                                      cfa(i) = .Fields(i)
                                      Combo7.AddItem cfa(i)
                                  End If
                              Next
                          End With
                      End If
                  End If
                  '案件備註
                  Text5 = pa(58)
                  
                  'Add By Sindy 2017/3/6 以內商程序身份進入時,改為巨京查名郵件
                  'Y52269.北京巨京知識產權代理有限公司
                  'If m_StrUserSt03 = "P22" And pa(10) = "020" And Left(m_strCP44, 6) = "Y52269" Then
                  If m_StrUserST03 = "P22" And pa(10) = "020" Then
                     If Left(m_strCP44, 6) <> "Y52269" Then '不是巨京,查詢巨京代表信箱
                        strExc(0) = "SELECT FA16 FROM FAGENT WHERE FA01='Y5226900' AND FA02='0'"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           m_strCP44 = "Y52269000" 'Add By Sindy 2018/5/11
                           m_Y52269FA16Mail = "" & RsTemp.Fields("FA16") '巨京代表信箱
                        End If
                     End If
                     '讀取最後收文承辦人
                     strCP14 = "": strCP14ST17 = ""
                     strSql = "SELECT cp14,st17 FROM CaseProgress,staff WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP09<'C' AND CP14 is not null AND CP14=ST01(+) ORDER BY cp05 desc,cp09 desc"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                     If intI = 1 Then
                        strCP14 = RsTemp.Fields("cp14")
                        strCP14ST17 = "" & RsTemp.Fields("st17")
                     End If
                     'Add By Sindy 2018/5/11
                     lblSendMailDt.Visible = True
                     lblSendMailDt.Caption = "寄件日期:"
                     '2018/5/11 END
                     cmdFCMail(0).Enabled = True
                  End If
                  '2017/3/6 END
              End If
               'add by sonia 2019/11/20
               bolDateType = DateType(pa(10), pa(23), pa(44))
               'end 2019/11/20
          Else
JumpToExitTM: 'Added by Lydia 2019/11/01
              Text1.SetFocus
              Text1_GotFocus
              'Added by Morgan 2013/3/8
              '案號錯誤時應跳離否則否可能會重複觸發事件導致無窮回圈
               Exit Sub
          End If
         
         '2008/9/18 ADD BY SONIA
         Combo8.Clear
         Combo8.AddItem "一般格式"
         'Modified by Morgan 2015/11/13 改單純下拉
         'Combo8.Text = "一般格式"
         Combo8.ListIndex = 0
         'end 2015/11/13
         
         Select Case Text1.Text
         'Add By Sindy 2009/07/02
         Case "TF"
            '核駁
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1002' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "核駁　　　　　　1002"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '2009/07/02 End
            'Add By Sindy 2009/08/21
            '被異議（理由）
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1602' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "被異議（理由）　1602"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '被評定（理由）
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1604' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "被評定（理由）　1604"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '被廢止（理由）
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1606' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "被廢止（理由）　1606"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '2009/08/21 End
            'add by sonia 2019/5/27
            '被部分廢止（理由）
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1620' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "被部分廢止（理由）1620"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
         Case "T"
            '核駁
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                      " AND CP10='1002' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1002")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "核駁　　　　　　1002"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '核駁先行通知
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1202' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1202")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "核駁前先行通知　1202"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '勝訴
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1003' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1003")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "勝訴　　　　　　1003"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '敗訴
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1004' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1004")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "敗訴　　　　　　1004"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            'add by sonia 2019/11/19
            '部分勝部分敗T-217896
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1006' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1006")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "部分勝部分敗　　1006"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            'end 2019/11/19
            '部分核駁(台->大)
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1205' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1205")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "部分核駁　　　　1205"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '2008/10/8 ADD BY SONIA
            '通知準備程序
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1203' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1203")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "通知準備程序　　1203"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '通知言詞辯論
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1204' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1204")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "通知言詞辯論　　1204"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '2008/10/17 ADD BY SONIA T-150838
            '撤銷原處分
            StrSQLa = "Select CP09,CP24 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1402' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1402", "T.CP09,T.CP24")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               If rsA.Fields(1) = "1" Then
                  Combo8.AddItem "撤銷原處分－勝　1402"
               Else
                  Combo8.AddItem "撤銷原處分－敗　1402"
               End If
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '2008/10/17 END
            '通知行政上訴答辯
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1406' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1406")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "通知行政上訴答辯1406"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '被異議
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1601' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1601")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "被異議　　　　　1601"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '被異議（理由）
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1602' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1602")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "被異議（理由）　1602"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '被評定
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1603' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1603")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "被評定　　　　　1603"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '被評定（理由）
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1604' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1604")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "被評定（理由）　1604"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '被廢止
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1605' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1605")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "被廢止　　　　　1605"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '被廢止（理由）
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1606' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1606")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "被廢止（理由）　1606"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '對方補充理由
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1609' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1609")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "對方補充理由　　1609"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '2008/10/20 ADD BY SONIA T-159978
            '補證據
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1617' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1617")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "補證據　　　　　1617"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '2008/10/20 END
            '對方答辯
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1618' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1618")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "對方答辯　　　　1618"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '智慧局答辯函
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1709' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1709")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "智慧局答辯函　　1709"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '2008/10/8 END
            
'            'Add By Sindy 2009/10/26
'            '變更申請案號
'            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'                            " AND CP10='1718' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' "
'            rsA.CursorLocation = adUseClient
'            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'            If rsA.RecordCount > 0 Then
'               Combo8.AddItem "變更申請案號　　1718"
'            End If
'            If rsA.State <> adStateClosed Then rsA.Close
'            Set rsA = Nothing
'            '2009/10/26 End
            
            '2009/2/26 add by sonia T-139747發回補答辯
            '發回補答辯
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1613' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1613")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "發回補答辯　　　1613"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '2009/2/26 end
            'Add By Sindy 2009/07/21
            '審查報告
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1201' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1201")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "審查報告　　　　1201"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '2009/07/21 End
            '2010/11/5 ADD BY SONIA  T-171559
            '通知修正
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1702' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1702")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "通知修正　　　　1702"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '2010/11/5 END
            'Add By Sindy 2009/08/12
            '通知復審答辯
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1404' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1404")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "通知復審答辯　　1404"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '2009/08/12 End
            '2012/4/27 add by sonia
            '對方撤回
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1610' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1610")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "對方撤回　　　　1610"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '2012/4/27 end
            'add by sonia 2019/5/27
            '被部分廢止
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1619' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1619")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "被部分廢止　　　1619"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '被部分廢止（理由）
            StrSQLa = "Select CP09 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1620' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' " & GetT727Sql("1620")
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Combo8.AddItem "被部分廢止（理由）1620"
               Combo8.ItemData(Combo8.ListCount - 1) = PUB_DocNo2Num(rsA(0)) '收文號(用來判斷是否印掛號) Added by Morgan 2015/11/12
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            'end 2019/5/27
         End Select
         '2008/9/18 END
         
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/30 +ACS系統類別
      Case "L", "FCL", "CFL", "LIN", "ACS"  '法務
          'edit by nickc 2007/02/02 不用 dll 了
          'If objPublicData.ReadLawCaseDatabase(pA()) Then
          If ClsPDReadLawCaseDatabase(pa()) Then
            m_bolReadOK = True 'Added by Morgan 2013/3/8
              If Not IsNull(pa()) Then
                  'Added by Lydia 2019/11/01 利益衝突案件：逐案號判斷
                  If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
                      If PUB_ChkCufaByCase(Me.Name, pa(1), pa(1) & pa(2) & pa(3) & pa(4), "" & pa(11) & "," & pa(43) & "," & pa(44) & "," & pa(45) & "," & pa(46), "" & pa(22)) = False Then
                          MsgBox MsgText(1109), vbInformation, MsgText(1110)
                          GoTo JumpToExitLC
                      End If
                  End If
                  'end 2019/11/01
                  
                  '案件名稱
                  If IsNull(pa(5)) = False And pa(5) <> "" Then Combo1.AddItem pa(5)
                  If IsNull(pa(6)) = False And pa(6) <> "" Then Combo1.AddItem pa(6)
                  If IsNull(pa(7)) = False And pa(7) <> "" Then Combo1.AddItem pa(7)
                  '2008/6/25 add by sonia
                  If IsNull(pa(17)) = False And pa(17) <> "" Then casetype4 = "(" & pa(17) & ")"
                  '2008/6/25 end
                  '2008/9/18 ADD BY SONIA
                  If IsNull(pa(16)) = False And pa(16) <> "" Then casetype5 = pa(16)
                  '2008/9/18 END
                  'Add By Cheng 2002/04/29
                  '是否閉卷
                  If Len("" & pa(8)) <= 0 Then
                      lblClose.Caption = ""
                  Else
                      lblClose.Caption = "已閉卷"
                  End If
                  Combo1 = pa(5)
                  '抓客戶基本檔
                  'Modify by Morgan 2011/9/16 +CU79
                  strExc(0) = "SELECT CU04,CU05,CU88,CU89,CU90,CU06,CU79 FROM CUSTOMER WHERE " & ChgCustomer(pa(11))
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                      With RsTemp
                          For i = 0 To 5
                              If IsNull(.Fields(i)) = False And (.Fields(i)) <> "" Then
                                  cu(i) = .Fields(i)
                                  Combo3.AddItem cu(i)
                              End If
                          Next
                          Text6(1) = "" & .Fields("CU79")
                      End With
                  End If

                  'Modified by Lydia 2020/06/04 備註和各項指示並存(暫時到完全取代), 原本是抓各項指示取代客戶備註
                  strMemoX = "": strKeyX = ""
                  If strSrvDate(1) >= 各項指示啟用日 Then
                      'Added by Lydia 2020/08/26 各項指示：完成確認，各項指示取代原先備註
                      If PUB_GetInstConfirm(m_StrUserST03, pa(11)) = True Then
                           Text6(1) = ""
                      End If
                      'end 2020/08/26
                      If Pub_GetInstructions(Me.Name, pa(11), strMemoX, , , , m_strIT10) Then
                          If strMemoX <> "" Then
                             Text6(1) = Text6(1) & IIf(Text6(1) <> "", vbCrLf, "") & strMemoX
                             strKeyX = pa(11)
                          End If
                      End If
                  End If
                  'end 2020/06/04
                 
                  '抓FCAgent
                  strExc(0) = "SELECT FA04,FA05,FA63,FA64,FA65,FA06 FROM FAGENT WHERE " & ChgFagent(pa(22))
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                      With RsTemp
                          For i = 0 To 5
                              If IsNull(.Fields(i)) = False And (.Fields(i)) <> "" Then
                                  fa(i) = .Fields(i)
                                  Combo2.AddItem fa(i)
                              End If
                          Next
                      End With
                  End If
                  'Add By Cheng 2003/06/26
                  '抓CFAgent
                  'Modify by Morgan 2004/10/7
                  'If bolFNation = True Then
                  If bolFA = True Then
                      '2006/5/23 MODIFY BY SONIA 加 CP09<'C'條件
                      '2009/6/23 MODIFY BY SONIA 郭雅娟說取消發文日CP27>0限制,因大陸已輸指示信未發文前無法使用
                      'Modify by Morgan 2011/1/5 +CP57 is null
                      StrSQLa = "Select CP44 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                                      " AND CP44 Is Not Null AND CP09<'C' and cp57 is null Order By CP27 Desc, CP09 Desc "
                      rsA.CursorLocation = adUseClient
                      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                      If rsA.RecordCount > 0 Then
                          m_strCP44 = "" & rsA.Fields(0).Value
                      Else
                          m_strCP44 = ""
                      End If
                      If rsA.State <> adStateClosed Then rsA.Close
                      Set rsA = Nothing
                      strExc(0) = "SELECT FA04,FA05,FA63,FA64,FA65,FA06 FROM FAGENT WHERE " & ChgFagent(m_strCP44)
                      '抓CFagent
                      intI = 1
                      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                      If intI = 1 Then
                          With RsTemp
                              For i = 0 To 5
                                  If IsNull(.Fields(i)) = False And (.Fields(i)) <> "" Then
                                      cfa(i) = .Fields(i)
                                      Combo7.AddItem cfa(i)
                                  End If
                              Next
                          End With
                      End If
                  End If
                  '案件備註
                  Text5 = pa(27)
              End If
               'add by sonia 2019/11/20
               bolDateType = DateType(pa(15), pa(11), pa(22))
               'end 2019/11/20
          Else
JumpToExitLC: 'Added by Lydia 2019/11/01
              Text1.SetFocus
              Text1_GotFocus
              'Added by Morgan 2013/3/8
              '案號錯誤時應跳離否則否可能會重複觸發事件導致無窮回圈
               Exit Sub
          End If
      
      Case "LA" '顧問
          'edit by nickc 2007/02/02 不用 dll 了
          'If objPublicData.ReadHireCaseDatabase(pA()) Then
          If ClsPDReadHireCaseDatabase(pa()) Then
            m_bolReadOK = True 'Added by Morgan 2013/3/8
              If Not IsNull(pa()) Then
                  'Added by Lydia 2019/11/01 利益衝突案件：逐案號判斷
                  If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
                      If PUB_ChkCufaByCase(Me.Name, pa(1), pa(1) & pa(2) & pa(3) & pa(4), "" & pa(5) & "," & pa(24) & "," & pa(25) & "," & pa(26) & "," & pa(27), "") = False Then
                          MsgBox MsgText(1109), vbInformation, MsgText(1110)
                          GoTo JumpToExitHC
                      End If
                  End If
                  'end 2019/11/01
                  
                  '案件名稱
                  If IsNull(pa(6)) = False And pa(6) <> "" Then Combo1.AddItem pa(6)
                  'Add By Cheng 2002/04/29
                  '2008/9/18 ADD BY SONIA
                  If IsNull(pa(7)) = False And pa(7) <> "" Then casetype5 = pa(7)
                  '2008/9/18 END
                  '是否閉卷
                  If Len("" & pa(9)) <= 0 Then
                      lblClose.Caption = ""
                  Else
                      lblClose.Caption = "已閉卷"
                  End If
                  Combo1 = pa(6)
                  '抓客戶基本檔
                  'Modify by Morgan 2011/9/16 +CU79
                  strExc(0) = "SELECT CU04,CU05,CU88,CU89,CU90,CU06,CU79 FROM CUSTOMER WHERE " & ChgCustomer(pa(5))
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                      With RsTemp
                          For i = 0 To 5
                              If IsNull(.Fields(i)) = False And (.Fields(i)) <> "" Then
                                  cu(i) = .Fields(i)
                                  Combo3.AddItem cu(i)
                              End If
                          Next
                          Text6(1) = "" & .Fields("CU79")
                      End With
                  End If
                  
                  'Modified by Lydia 2020/06/04 備註和各項指示並存(暫時到完全取代), 原本是抓各項指示取代客戶備註
                  strMemoX = "": strKeyX = ""
                  If strSrvDate(1) >= 各項指示啟用日 Then
                      'Added by Lydia 2020/08/26 各項指示：完成確認，各項指示取代原先備註
                      If PUB_GetInstConfirm(m_StrUserST03, pa(5)) = True Then
                           Text6(1) = ""
                      End If
                      'end 2020/08/26
                      If Pub_GetInstructions(Me.Name, pa(5), strMemoX, , , , m_strIT10) Then
                          If strMemoX <> "" Then
                             Text6(1) = Text6(1) & IIf(Text6(1) <> "", vbCrLf, "") & strMemoX
                             strKeyX = pa(5)
                          End If
                      End If
                  End If
                  'end 2020/06/04
                  
                  '案件備註
                  Text5 = pa(12)
              End If
               'add by sonia 2019/11/20
               bolDateType = DateType("000", pa(5), "")
               'end 2019/11/20
          Else
JumpToExitHC: 'Added by Lydia 2019/11/01
              Text1.SetFocus
              Text1_GotFocus
              'Added by Morgan 2013/3/8
              '案號錯誤時應跳離否則否可能會重複觸發事件導致無窮回圈
               Exit Sub
          End If
      
      Case Else '服務業務
          'edit by nickc 2007/02/02 不用 dll 了
          'If objPublicData.ReadServicePracticeDatabase(pA(), intWhere) Then
          If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then
            m_bolReadOK = True 'Added by Morgan 2013/3/8
              If Not IsNull(pa()) Then
                  'Added by Lydia 2019/11/01 利益衝突案件：逐案號判斷
                  If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
                      If PUB_ChkCufaByCase(Me.Name, pa(1), pa(1) & pa(2) & pa(3) & pa(4), "" & pa(8) & "," & pa(58) & "," & pa(59) & "," & pa(65) & "," & pa(66), "" & pa(26)) = False Then
                          MsgBox MsgText(1109), vbInformation, MsgText(1110)
                          GoTo JumpToExitSP
                      End If
                  End If
                  'end 2019/11/01
                  
                  '案件名稱
                  If IsNull(pa(5)) = False And pa(5) <> "" Then Combo1.AddItem pa(5)
                  If IsNull(pa(6)) = False And pa(6) <> "" Then Combo1.AddItem pa(6)
                  If IsNull(pa(7)) = False And pa(7) <> "" Then Combo1.AddItem pa(7)
                  '2008/6/25 add by sonia
                  If IsNull(pa(29)) = False And pa(29) <> "" Then casetype4 = "(" & pa(29) & ")"
                  '2008/6/25 end
                  '2008/9/18 ADD BY SONIA
                  If IsNull(pa(28)) = False And pa(28) <> "" Then casetype5 = pa(28)
                  '2008/9/18 END
                  'Add By Cheng 2002/04/29
                  '是否閉卷
                  If Len("" & pa(15)) <= 0 Then
                      lblClose.Caption = ""
                  Else
                      lblClose.Caption = "已閉卷"
                  End If
                  '申請國家1
                  Label11 = pa(9)
                  'edit by nickc 2007/02/02 不用 dll 了
                  'If objPublicData.GetNation(pA(9), strTemp) Then
                  If ClsPDGetNation(pa(9), strTemp) Then
                      Label12.Caption = strTemp
                  End If
                  Combo1 = pa(5)
                  '抓客戶基本檔
                  'Modify By Sindy 2009/08/13
                  'strExc(0) = "SELECT CU04,CU05,CU88,CU89,CU90,CU06 FROM CUSTOMER WHERE " & ChgCustomer(pa(26))
                  'Modify by Morgan 2011/9/16 +CU79
                  strExc(0) = "SELECT CU04,CU05,CU88,CU89,CU90,CU06,CU79 FROM CUSTOMER WHERE " & ChgCustomer(pa(8))
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                      With RsTemp
                          For i = 0 To 5
                              If IsNull(.Fields(i)) = False And (.Fields(i)) <> "" Then
                                  cu(i) = .Fields(i)
                                  Combo3.AddItem cu(i)
                              End If
                          Next
                          Text6(1) = "" & .Fields("CU79")
                      End With
                  End If

                  'Modified by Lydia 2020/06/04 備註和各項指示並存(暫時到完全取代), 原本是抓各項指示取代客戶備註
                  strMemoX = "": strKeyX = ""
                  If strSrvDate(1) >= 各項指示啟用日 Then
                      'Added by Lydia 2020/08/26 各項指示：完成確認，各項指示取代原先備註
                      If PUB_GetInstConfirm(m_StrUserST03, pa(8)) = True Then
                           Text6(1) = ""
                      End If
                      'end 2020/08/26
                      If Pub_GetInstructions(Me.Name, pa(8), strMemoX, , , , m_strIT10) Then
                          If strMemoX <> "" Then
                             Text6(1) = Text6(1) & IIf(Text6(1) <> "", vbCrLf, "") & strMemoX
                             strKeyX = pa(8)
                          End If
                      End If
                  End If
                  'end 2020/06/04
                  
                  '抓FCAgent
                  strExc(0) = "SELECT FA04,FA05,FA63,FA64,FA65,FA06 FROM FAGENT WHERE " & ChgFagent(pa(26))
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                      With RsTemp
                          For i = 0 To 5
                              If IsNull(.Fields(i)) = False And (.Fields(i)) <> "" Then
                                  fa(i) = .Fields(i)
                                  Combo2.AddItem fa(i)
                              End If
                          Next
                      End With
                  End If
                   
                  '抓CFAgent
                  If bolFA = True Then
                      '2006/5/23 MODIFY BY SONIA 加 CP09<'C'條件
                      '2009/6/23 MODIFY BY SONIA 郭雅娟說取消發文日CP27>0限制,因大陸已輸指示信未發文前無法使用
                      'Modify by Morgan 2011/1/5 +CP57 is null
                      StrSQLa = "Select CP44 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                                      " AND CP44 Is Not Null AND CP09<'C' and cp57 is null Order By CP27 Desc, CP09 Desc "
                      rsA.CursorLocation = adUseClient
                      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                      If rsA.RecordCount > 0 Then
                          m_strCP44 = "" & rsA.Fields(0).Value
                      Else
                          m_strCP44 = ""
                      End If
                      If rsA.State <> adStateClosed Then rsA.Close
                      Set rsA = Nothing
                      strExc(0) = "SELECT FA04,FA05,FA63,FA64,FA65,FA06 FROM FAGENT WHERE " & ChgFagent(m_strCP44)
                      '抓CFagent
                      intI = 1
                      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                      If intI = 1 Then
                          With RsTemp
                              For i = 0 To 5
                                  If IsNull(.Fields(i)) = False And (.Fields(i)) <> "" Then
                                      cfa(i) = .Fields(i)
                                      Combo7.AddItem cfa(i)
                                  End If
                              Next
                          End With
                      End If
                  End If
                  '案件進度
                  Text5 = pa(18)
              End If
               'add by sonia 2019/11/20
               bolDateType = DateType(pa(9), pa(8), pa(26))
               'end 2019/11/20
          Else
JumpToExitSP: 'Added by Lydia 2019/11/01
              Text1.SetFocus
              Text1_GotFocus
              'Added by Morgan 2013/3/8
              '案號錯誤時應跳離否則否可能會重複觸發事件導致無窮回圈
               Exit Sub
          End If
   End Select
  
    'Modified by Lydia 2020/06/04 備註和各項指示並存(暫時到完全取代), 原本是抓各項指示取代客戶備註
    strMemoCase = "": strKeyCase = ""
    If strSrvDate(1) >= 各項指示啟用日 Then
        'Added by Lydia 2020/08/26 各項指示：完成確認，各項指示取代原先備註
        If PUB_GetInstConfirm(m_StrUserST03, pa(1) & pa(2) & pa(3) & pa(4)) = True Then
             Text5 = ""
        End If
        'end 2020/08/26
        If Pub_GetInstructions(Me.Name, pa(1) & pa(2) & pa(3) & pa(4), strMemoCase, , , , m_strIT10) Then
           If strMemoCase <> "" Then
               Text5 = Text5 & IIf(Text5 <> "", vbCrLf, "") & strMemoCase
               strKeyCase = pa(1) & pa(2) & pa(3) & pa(4)
           End If
        End If
    End If
    'end 2020/06/04
    
   If pa(1) <> "FCP" Then 'Add by Morgan 2007/3/12 FCP已有預設
      If Combo1.ListCount > 0 Then Combo1.ListIndex = 0 'Added by Lydia 2017/04/06 預設在第一個名稱
      If Combo2.ListCount > 0 Then Combo2.ListIndex = 0 'Add by Morgan 2004/9/16
      If Combo3.ListCount > 0 Then Combo3.ListIndex = 0 'Add by Morgan 2004/9/16
      If Combo7.ListCount > 0 Then Combo7.ListIndex = 0 'Add by Morgan 2006/10/26
      'Add By Sindy 2015/1/21 FG要同FCP預設值 ex.FG-000997
      If pa(1) = "FG" Then
         If Combo2.List(0) <> "" Then
            If m_iLang = -1 Then m_iLang = 1
            If Option1(m_iLang).Value = True Then
               Option1_Click m_iLang
            Else
               Option1(m_iLang).Value = True
            End If
            If Option2.Value = True Then
               Option2_Click
            Else
               Option2.Value = True
            End If
            Option4.Value = True
         End If
      End If
      '2015/1/21 END
   End If
   m_bEMail = PUB_GetEMailFlag(pa(1) & pa(2) & pa(3) & pa(4))
   
   'Add by Morgan 2010/12/14 業務預設中文,申請人
   If Left(Pub_StrUserSt15, 1) = "S" Then
      If (Option1(0).Value Or Option1(1).Value Or Option1(2).Value) = False Then
         Option1(0).Value = True
      End If
      Option3.Value = 1
   End If
End Sub

'T台灣案分析相關總收文號
Private Function GetT727Sql(strCP10 As String, Optional strCols As String = "T.CP09") As String
   GetT727Sql = " union Select " & strCols & " FROM CaseProgress C,(" & _
                " Select CP43,CP09,CP24,CP27,CP49 FROM CaseProgress" & _
                " WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10='727'" & _
                " AND CP27 IS NULL AND CP57 IS NULL) T" & _
                " Where c.CP09 = t.CP43 and c.cp10='" & strCP10 & "'"
End Function

Private Sub WordChinese_sub1(g_WordAp As Word.Application)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim m_CP07 As String    '法定期限/中民
Dim m_CP06 As String    '本所期限/中民
Dim m_TIME As String    '開庭時間2008/11/12 ADD BY SONIA
Dim m_PLACE As String   '開庭地點2008/11/12 ADD BY SONIA
Dim i As String         '2008/11/12 ADD BY SONIA
Dim m_DueDT As String
Dim m_DueDate As String
Dim m_CPM04 As String    '下一程序   'ADD BY SONIA 2015/4/16
Dim m_0111006 As Boolean  '日本最終核駁  add by sonia 2015/9/1 CFP-025108
Dim strFirstPriDate As String  '最早的優先權日期  2009/3/3 ADD BY SONIA
Dim strRefCP10 As String 'Added by Morgan 2015/5/26
Dim m_AppDate As String 'Added by Morgan 2021/9/3 約定期限
Dim strCP36Kind As String
Dim tmpArr As Variant

m_AppDate = "　　年　　月　　日" 'Added by Morgan 2021/9/3

      '2008/6/25 MODIFY BY SONIA 分信函格式
      '2008/9/18 MODIFY BY SONIA 加入內商格式
      Select Case Left(m_Combo8, 1)
         Case "0"
            Select Case m_Combo8
               Case "00"      '一般格式
                  '2008/9/19 ADD BY SONIA
                  If Text1.Text = "T" Or Text1.Text = "TF" Or Text1.Text = "CFT" Or Text1.Text = "FCT" Then
                     casetype2 = ""
                  End If
                  '2008/9/19 END
                  '2008/12/1 MODIFY BY SONIA 加申請案號
                  'g_WordAp.Selection.TypeText "　　" & custtype & "委由本所辦理之「" & CASENAME & "」" & casetype4 & nationname & casetype1 & casetype2 & casetype3 & "案(本所案號" & caseno & ")，頃接"
                  '2008/12/5 modify by sonia 商標案加商品類別
                  'g_WordAp.Selection.TypeText "　　" & custtype & "委由本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & nationname & casetype1 & casetype2 & casetype3 & "案(本所案號" & caseno & ")，頃接"
                  'Modified by Lydia 2020/10/15 委由=>委託
                  If Text1.Text = "T" Or Text1.Text = "TF" Or Text1.Text = "CFT" Or Text1.Text = "FCT" Then
                     'Modified by Lydia 2020/11/02 改描述
                     'g_WordAp.Selection.TypeText "　　" & custtype & "委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & nationname & "第" & pa(9) & "類" & casetype1 & casetype2 & casetype3 & "案(本所案號" & CaseNo & ")，頃接"
                     strExc(3) = "　　" & custtype & "委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & nationname & _
                                   "第" & pa(9) & "類" & casetype1 & casetype2 & casetype3 & "案，頃接......隨函附送......，敬請查收備存。"
                     g_WordAp.Selection.TypeText strExc(3)
                  Else
                     'Modified by Lydia 2020/11/02 改描述
                     'g_WordAp.Selection.TypeText "　　" & custtype & "委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & nationname & casetype1 & casetype2 & casetype3 & "案(本所案號" & CaseNo & ")，頃接"
                     strExc(3) = "　　" & custtype & "委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & nationname & _
                                   casetype1 & casetype2 & casetype3 & "案，頃接......隨函附送......，敬請查收備存。"
                     g_WordAp.Selection.TypeText strExc(3)
                  End If
                  '2008/12/5 end
                  '2008/12/1 END
                  'Added by Lydia 2020/11/02 加描述
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　本案本所案號為「" & CaseNo & "」，往後來函查詢時，請註明本所案號，以利處理。"
                  'end 2020/11/02
                  g_WordAp.Selection.TypeParagraph
                  '2008/5/1 end
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  'Added by Lydia 2020/11/02 加描述
                  g_WordAp.Selection.TypeText "(致客戶之敬語請依案情斟酌選擇，完稿時刪除不適用的敬語及說明文字。)"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "(第一種型態針對提出向官方提出申請：)"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　關於本案之進度，本所當隨時與　" & custtype & "聯繫，如有任何問題請隨時不吝賜教，本所將竭誠提供服務。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "(第二種型態為已完成客戶交辦之事務：)"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　如有任何問題請隨時不吝賜教，本所將竭誠提供服務。"
                  'end 2020/11/02
                  g_WordAp.Selection.TypeParagraph
                  'Added by Morgan 2015/5/14
                  'Modified by Morgan 2021/10/19 +未發文的OA也同分析出通用定稿
                  'If Combo8 = "分析　　　　　　941 " And m_AboutDeadLine <> "" Then
                  If Right(Combo8, 1) = " " And m_AboutDeadLine <> "" Then
                  'end 2021/10/19
                     g_WordAp.Selection.TypeText Replace(m_AboutDeadLine, "貴單位", custtype)
                  End If
                  'end 2015/5/14
                  
                  g_WordAp.Selection.TypeParagraph
                  'Remove by Lydia 2020/11/02 因為上面加註
                  'g_WordAp.Selection.TypeParagraph
                  'g_WordAp.Selection.TypeParagraph
                  'g_WordAp.Selection.TypeParagraph
                  'g_WordAp.Selection.TypeParagraph
                  'g_WordAp.Selection.TypeParagraph
                  'end 2020/11/02
                  '2008/5/1 add by sonia 專利處要求加印
                  g_WordAp.Selection.TypeParagraph
                  'g_WordAp.Selection.TypeText "　　其他若有任何質疑，亦請不吝賜教，本所當竭誠提供最佳服務。"　'Remove by Lydia 2020/11/02
            End Select
         Case "1"
            Select Case m_Combo8
               Case "11"      'CFP核駁
                  '2011/3/16 modify by sonia 非美國之最終核駁也用此,下一程序抓資料庫不寫死
                  'StrSQLa = "Select CP07 FROM CaseProgress WHERE " & ChgCaseProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1002' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' "
                  'Modified by Morgan 2021/9/3 +NP23及NP+本所案號條件否則可能會抓到其他案號來的IDS的期限
                  'StrSQLa = "Select CP07,CPM04,CP10 FROM CaseProgress,NEXTPROGRESS,CASEPROPERTYMAP WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10 in ('1002','1006') AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' AND CP09=NP01(+) AND NP02=CPM01(+) AND NP07=CPM02(+) "
                  StrSQLa = "Select CP07,CPM04,CP10,NP23 FROM CaseProgress,NEXTPROGRESS,CASEPROPERTYMAP WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10 in ('1002','1006') AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' AND CP09=NP01(+) AND CP01=NP02(+) AND CP02=NP03(+) AND CP03=NP04(+) AND CP04=NP05(+) AND NP02=CPM01(+) AND NP07=CPM02(+) "
                  '2011/3/16 end
                  '2013/1/21 ADD BY SONIA 加入分析941同核駁1002格式
                  'Modified by Morgan 2021/9/3 +NP23(怪怪的,分析應該不會有下一程序)
                  'Modified by Morgan 2021/10/6 分析應該用相關收文號抓下一程序才對
                  StrSQLa = StrSQLa & "UNION Select NP09,CPM04,CP10,NP23 FROM CaseProgress,NEXTPROGRESS,CASEPROPERTYMAP WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='941' AND CP27 IS NULL AND CP57 IS NULL AND CP09<'C' AND CP43=NP01(+) AND NP02=CPM01(+) AND NP07=CPM02(+) "
                  '2013/1/21 END
                  
                  rsA.CursorLocation = adUseClient
                  rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                  If rsA.RecordCount > 0 Then
                     m_CP07 = Mid(rsA.Fields(0).Value, 1, 4) & "年" & Mid(rsA.Fields(0).Value, 5, 2) & "月" & Mid(rsA.Fields(0).Value, 7, 2) & "日"
                     m_CPM04 = "" & rsA.Fields(1).Value   'ADD BY SONIA 2015/4/16
                     'Added by Morgan 2021/9/3 約定期限
                     If rsA("NP23") > 0 Then
                        m_AppDate = Mid(rsA("NP23"), 1, 4) & "年" & Mid(rsA("NP23"), 5, 2) & "月" & Mid(rsA("NP23"), 7, 2) & "日"
                     End If
                     'end 2021/9/3
                  End If
                  
                  'add by sonia 2015/9/1 日本011最終核駁1006要加一句話CFP-025108
                  m_0111006 = False
                  If pa(9) = "011" And rsA.Fields(2).Value = "1006" Then m_0111006 = True
                  'end 2015/9/1
                  'Modified by Morgan 2018/1/3 EPC專利局->EPO --玫音
                  'Modified by Morgan 2024/5/16 核駁審定書->審查意見通知書,EPO改半形 --郭
                  'g_WordAp.Selection.TypeText "　　一、　" & custtype & "前委託本所代理向" & nationname & "申請之" & AppNo & "「" & CASENAME & "」" & casetype4 & casetype1 & casetype2 & casetype3 & "案(本所案號" & CaseNo & ")，頃接" & nationname & "專利代理人轉來" & IIf(nationname = "ＥＰＣ", "ＥＰＯ", nationname & "專利局") & "之核駁審定書，謂本案暫不准予專利。"
                  g_WordAp.Selection.TypeText "　　一、　" & custtype & "前委託本所代理向" & nationname & "申請之" & AppNo & "「" & CASENAME & "」" & casetype4 & casetype1 & casetype2 & casetype3 & "案(本所案號" & CaseNo & ")，頃接" & nationname & "專利代理人轉來" & IIf(pa(9) = "221", "EPO", nationname & "專利局") & "之審查意見通知書，謂本案暫不准予專利。"
                  'end 2024/5/16
                  g_WordAp.Selection.TypeParagraph
'                  'modify by sonia 2015/9/1 日本011最終核駁1006要加一句話CFP-025108
                  'Modified by Morgan 2024/5/16 審定書->審查意見通知書 --郭
                  g_WordAp.Selection.TypeText "　　二、茲隨函附上審查意見通知書與前案影本乙份，敬請存查。"
'                  If m_0111006 Then
'                     g_WordAp.Selection.TypeText "　　二、茲隨函附上審定書與前案影本乙份，同時附上提出訴願時需一併提交之委任狀，敬請存查。"
'                  Else
'                     g_WordAp.Selection.TypeText "　　二、茲隨函附上審定書與前案影本乙份，敬請存查。"
'                  End If
'                  'end 2015/9/1
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　三、審查意見："
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　四、本所意見："
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  If rsA.State <> adStateClosed Then rsA.Close
                  Set rsA = Nothing
                  '2011/3/16 modify by sonia
                  'g_WordAp.Selection.TypeText "　　五、本案依法應於" & m_CP07 & "以前提出答辯。本所代　" & custtype & "答辯所需之費用為新台幣　　萬　　仟元整。倘　" & custtype & "續行本案，敬請於　　年　　月　　日以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向該國專利局提出。若逾上述與本所聯絡的期限，除必要之延期費外，本所將另外增收服務費。事關　" & custtype & "權益，尚請特加留意是幸。"
                  '2015/4/16 MODIFY BY SONIA 依是否有美國關聯案決定要不要有 前案揭露聲明 的段落 CFP-027079, 且美國案依大小微個體之分,費用不同
                  'g_WordAp.Selection.TypeText "　　五、本案依法應於" & m_CP07 & "以前提出" & "" & rsA.Fields(1).Value & "。本所代　" & custtype & "答辯所需之費用為新台幣　　萬　　仟元整。倘　" & custtype & "續行本案，敬請於　　年　　月　　日以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向該國專利局提出。若逾上述與本所聯絡的期限，除必要之延期費外，本所將另外增收服務費。事關　" & custtype & "權益，尚請特加留意是幸。"
                     
                  'Modified by Morgan 2018/5/30
                  'StrSQLa = "SELECT CR05||'-'||CR06||DECODE(CR07,'0',NULL,'-'||CR07||'-'||CR08) CR101NO,DECODE(SUBSTR(PA91,INSTR(PA91,'個體')-1,1),'大','貳萬陸仟元','小','貳萬參仟元','微','貳萬壹仟元','') FEE FROM CASERELATION,PATENT" & _
                            " WHERE CR01='" & pa(1) & "' AND CR02='" & pa(2) & "' AND CR03='" & pa(3) & "' AND CR04='" & pa(4) & "'" & _
                            " AND CR05=PA01(+) AND CR06=PA02(+) AND CR07=PA03(+) AND CR08=PA04(+) AND PA09='101' "
                  '2013/4/26 end
                  'rsA.CursorLocation = adUseClient
                  'rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                  ''有美國關聯案
                  'If rsA.RecordCount > 0 Then
                  If fnUsIdsChk(rsA) = True Then
                  'end 2018/5/30
                        g_WordAp.Selection.Font.ColorIndex = wdRed
                        g_WordAp.Selection.TypeText "　　五、前案揭露聲明（INFORMATION DISCLOSURE STATEMENT, IDS）《註：請承辦工程師確認本節IDS內容是否留存。若與本案相關之美國專利案已取得專利證書，或本案的相關前案均已見於相關美國專利案，則本節IDS內容無須留存。請工程師閱畢後刪除本段註記。》"
                        g_WordAp.Selection.TypeParagraph
                        'modify by sonia 2016/3/14 慧汶請作單取消固定費用改個案報價
                        'modify by sonia 2016/4/12
                        'g_WordAp.Selection.TypeText "　　按依照美國專利法之規定，申請人有義務將所知與本案有關的前案主動提交美國專利局審查員作為審查與本案相關之美國專利案之參考，即美國專利法所規定之前案揭露聲明。若有明知而未陳報而被發現或者遭他人舉報者，將對相關美國專利案造成不利之後果。經查　" & custtype & "另申請有與本案相關之美國專利申請案（本所案號" & rsA.Fields("CR101NO") & "），因此依美國專利法之規定，應於三個月內主動將本次審查意見通知書（檢索報告）所附前案向美國專利局提出，以符合美國專利法之規定，避免對相關美國專利案產生不利的後果。另外美國提出前案揭露所需費用會依所提出時間不同而有所不同，若在該美國專利案第一次審定書發出前提出前案揭露聲明，本所代　" & custtype & "提出前案揭露聲明的費用為壹萬柒仟元整（經查該與本案相關之美國專利案目前並未收到第一次審定書），而若在發出第一次審定書後方才提出前案揭露聲明，則需費用" & rsA.Fields("FEE") & "整，在此一併告知　" & custtype & "，事關　" & custtype & "權益，尚請特加留意是幸。"
                        'Modified by Morgan 2019/12/23 內容調整--郭
                        'g_WordAp.Selection.TypeText "　　按依照美國專利法之規定，申請人有義務將所知與本案有關的前案主動提交美國專利局審查員作為審查與本案相關之美國專利案之參考，即美國專利法所規定之前案揭露聲明。若有明知而未陳報而被發現或者遭他人舉報者，將對相關美國專利案造成不利之後果。經查　" & custtype & "另申請有與本案相關之美國專利申請案（本所案號" & rsA.Fields("CR101NO") & "），因此依美國專利法之規定，應於三個月內主動將本次審查意見通知書（檢索報告）所附前案向美國專利局提出，以符合美國專利法之規定，避免對相關美國專利案產生不利的後果。另外美國提出前案揭露所需費用會依所提出時間不同而有所不同，若在該美國專利案第一次審定書發出前提出前案揭露聲明，本所代　" & custtype & "提出前案揭露聲明的費用為新台幣　　萬　　仟元整（經查該與本案相關之美國專利案目前並未收到第一次審定書），而若在發出第一次審定書後方才提出前案揭露聲明，則需費用新台幣　　萬　　仟元整，在此一併告知　" & custtype & "，事關　" & custtype & "權益，尚請特加留意是幸。"
                        'Modified by Morgan 2020/3/23 內容調整--郭
                        'g_WordAp.Selection.TypeText "　　按依照美國專利法之規定，申請人有義務將所知與本案有關的前案主動提交美國專利局審查員作為審查與本案相關之美國專利案之參考，即美國專利法所規定之前案揭露聲明。若有明知而未陳報而被發現或者遭他人舉報者，將對相關美國專利案造成不利之後果。經查　" & custtype & "另申請有與本案相關之美國專利申請案（本所案號" & rsA.Fields("CR101NO") & "），因此依美國專利法之規定，應於首次得知的三個月內主動將本次審查意見通知書（檢索報告）所附前案向美國專利局提出，以符合美國專利法之規定，避免對相關美國專利案產生不利的後果，此階段所需費用為新台幣　　萬　　仟元整；惟若是超過三個月後且美國專利申請案已發出審查意見通知書才提出前案揭露聲明者，此階段所需費用為新台幣　　萬　　仟元整，在此一併告知，事關　" & custtype & "權益，尚請特加留意是幸。"
                        'Modified by Morgan 2020/8/6 內容調整--郭
                        'g_WordAp.Selection.TypeText "　　按依照美國專利法之規定，申請人有義務將所知與本案有關的前案主動提交美國專利局審查員作為審查與本案相關之美國專利案之參考，即美國專利法所規定之前案揭露聲明。若有明知而未陳報而被發現或者遭他人舉報者，將對相關美國專利案造成不利之後果。經查　" & custtype & "另申請有與本案相關之美國專利申請案（本所案號" & rsA.Fields("CR101NO") & "），因此依美國專利法之規定，應於首次得知的三個月內主動將本次審查意見通知書（檢索報告）所附前案向美國專利局提出，以符合美國專利法之規定，避免對相關美國專利案產生不利的後果，此階段所需費用為新台幣　　萬　　仟元整；惟若是超過三個月後且美國專利申請案已發出審查意見通知書才提出前案揭露聲明者，此階段所需費用為新台幣　　萬　　仟元整；若是在收到美國專利申請案之最終審定後才提出前案揭露聲明者，費用另計，在此一併告知，事關　" & custtype & "權益，尚請特加留意是幸。"
                        'Modified by Morgan 2023/10/30 內容調整--郭
                        'g_WordAp.Selection.TypeText "　　按依照美國專利法之規定，申請人有義務將所知與本案有關的前案主動提交美國專利局審查員作為審查與本案相關之美國專利案之參考，即美國專利法所規定之前案揭露聲明。若有明知而未陳報而被發現或者遭他人舉報者，將對相關美國專利案造成不利之後果。經查　" & custtype & "另申請有與本案相關之美國專利申請案（本所案號" & rsA.Fields("CR101NO") & "），因此依美國專利法之規定，應於首次得知的三個月內主動將本次審查意見通知書（檢索報告）所附前案向美國專利局提出，以符合美國專利法之規定，避免對相關美國專利案產生不利的後果，此階段所需費用為新台幣　　萬　　仟元整；惟若是超過三個月後且美國專利申請案已發出審查意見通知書才提出前案揭露聲明者，或是美國專利申請案已接獲最終核駁或核准通知書(在未繳納領證費前)才提出前案揭露聲明者，此階段所需費用為新台幣　　萬　　仟元整，在此一併告知，事關　" & custtype & "權益，尚請特加留意是幸。"
                        'Modified by Morgan 2025/5/22 內容調整--郭
                        'g_WordAp.Selection.TypeText "　　按依照美國專利法之規定，申請人有義務將所知與本案有關的前案主動提交美國專利局審查員作為審查與本案相關之美國專利案之參考，即美國專利法所規定之前案揭露聲明。若有明知而未陳報而被發現或者遭他人舉報者，將對相關美國專利案造成不利之後果。經查　" & custtype & "另申請有與本案相關之美國專利申請案（本所案號" & rsA.Fields("CR101NO") & "），因此依美國專利法之規定，應於首次得知的三個月內主動將本次審查意見通知書（檢索報告）所附前案向美國專利局提出，以符合美國專利法之規定，避免對相關美國專利案產生不利的後果，此階段所需費用為新台幣　　萬　　仟元整(前案以25件為限，每超出5件須加收新台幣2,500元)；惟若是超過三個月後且美國專利申請案已發出審查意見通知書才提出前案揭露聲明者，或是美國專利申請案已接獲最終核駁或核准通知書(在未繳納領證費前)才提出前案揭露聲明者，此階段所需費用為新台幣　　萬　　仟元整(前案以25件為限，每超出5件須加收新台幣2,500元)，在此一併告知，事關　" & custtype & "權益，尚請特加留意是幸。"
                        g_WordAp.Selection.TypeText "　　按依照美國專利法之規定，申請人有義務將所知與本案有關的前案主動提交美國專利局審查員作為審查與本案相關之美國專利案之參考，即美國專利法所規定之前案揭露聲明。若有明知而未陳報而被發現或者遭他人舉報者，將對相關美國專利案造成不利之後果。經查　" & custtype & "另申請有與本案相關之美國專利申請案（本所案號" & rsA.Fields("CR101NO") & "），因此依美國專利法之規定，應於首次得知的三個月內主動將本次審查意見通知書（檢索報告）所附前案向美國專利局提出，以符合美國專利法之規定，避免對相關美國專利案產生不利的後果，此階段所需費用為新台幣　　萬　　仟元整(前案以25件為限，每超出5件須加收新台幣2,500元)；"
                        g_WordAp.Selection.TypeParagraph
                        g_WordAp.Selection.TypeText "　　惟若是超過三個月後且美國專利申請案已發出審查意見通知書才提出前案揭露聲明者，或是雖未超過三個月但美國專利申請案已接獲最終核駁或核准通知書(在未繳納領證費前)才提出前案揭露聲明者，此階段所需費用為新台幣　　萬　　仟元整(前案以25件為限，每超出5件須加收新台幣2,500元);"
                        g_WordAp.Selection.TypeParagraph
                        g_WordAp.Selection.TypeText "　　惟若是超過三個月後且美國專利申請案已接獲最終核駁或核准通知書(在未繳納領證費前)才提出前案揭露聲明者，此階段的程序較為複雜，將依個案情形另外提供報價；"
                        g_WordAp.Selection.TypeParagraph
                        g_WordAp.Selection.TypeText "　　在此一併告知，事關　" & custtype & "權益，尚請特加留意是幸。"
                        'end 2025/5/22
                        'end 2023/10/30
                        'end 2020/8/6
                        g_WordAp.Selection.Font.ColorIndex = wdAuto
                        g_WordAp.Selection.TypeParagraph
                        g_WordAp.Selection.TypeParagraph
                     'modify by sonia 2015/9/1 日本011最終核駁1006要加一句話CFP-025108
                     'g_WordAp.Selection.TypeText "　　六、本案依法應於" & m_CP07 & "以前提出" & m_CPM04 & "。本所代　" & custtype & "答辯所需之費用為新台幣　　萬　　仟元整。倘　" & custtype & "續行本案，敬請於　　年　　月　　日以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向該國專利局提出。若逾上述與本所聯絡的期限，除必要之延期費外，本所將另外增收服務費。事關　" & custtype & "權益，尚請特加留意是幸。"
                  'Modified yb Morgan 2015/9/18 本所代　" & custtype & "答辯所需之費用-->本所代　" & custtype & m_CPM04 & "所需之費用 --說法改一致--偉城,郭
                     If m_0111006 Then
                        'Modified by Morgan 2021/9/3
                        'g_WordAp.Selection.TypeText "　　六、本案依法應於" & m_CP07 & "以前提出" & m_CPM04 & "。本所代　" & custtype & m_CPM04 & "所需之費用為新台幣　　萬　　仟元整。倘　" & custtype & "續行本案，敬請於　　年　　月　　日以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向該國專利局提出。若逾上述與本所聯絡的期限，除必要之延期費外，本所將另外增收服務費。事關　" & custtype & "權益，尚請特加留意是幸。"
                        g_WordAp.Selection.TypeText "　　六、本案依法應於" & m_CP07 & "以前提出" & m_CPM04 & "。本所代　" & custtype & m_CPM04 & "所需之費用為新台幣　　萬　　仟元整。倘　" & custtype & "續行本案，敬請於" & m_AppDate & "以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向該國專利局提出。若逾上述與本所聯絡的期限，除必要之延期費外，本所將另外增收服務費，另外，本所服務費也會因匯率變動關係而有所調整。事關　" & custtype & "權益，尚請特加留意是幸。"
                     Else
                        'Modified by Morgan 2021/9/3
                        'g_WordAp.Selection.TypeText "　　六、本案依法應於" & m_CP07 & "以前提出" & m_CPM04 & "。本所代　" & custtype & m_CPM04 & "所需之費用為新台幣　　萬　　仟元整。倘　" & custtype & "續行本案，敬請於　　年　　月　　日以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向該國專利局提出。若逾上述與本所聯絡的期限，除必要之延期費外，本所將另外增收服務費。事關　" & custtype & "權益，尚請特加留意是幸。"
                        g_WordAp.Selection.TypeText "　　六、本案依法應於" & m_CP07 & "以前提出" & m_CPM04 & "。本所代　" & custtype & m_CPM04 & "所需之費用為新台幣　　萬　　仟元整。倘　" & custtype & "續行本案，敬請於" & m_AppDate & "以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向該國專利局提出。若逾上述與本所聯絡的期限，除必要之延期費外，本所將另外增收服務費，另外，本所服務費也會因匯率變動關係而有所調整。事關　" & custtype & "權益，尚請特加留意是幸。"
                     End If
                  Else
                     'modify by sonia 2015/9/1 日本011最終核駁1006要加一句話CFP-025108
                     'g_WordAp.Selection.TypeText "　　五、本案依法應於" & m_CP07 & "以前提出" & m_CPM04 & "。本所代　" & custtype & "答辯所需之費用為新台幣　　萬　　仟元整。倘　" & custtype & "續行本案，敬請於　　年　　月　　日以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向該國專利局提出。若逾上述與本所聯絡的期限，除必要之延期費外，本所將另外增收服務費。事關　" & custtype & "權益，尚請特加留意是幸。"
                     If m_0111006 Then
                        'Modified by Morgan 2021/9/3
                        'g_WordAp.Selection.TypeText "　　五、本案依法應於" & m_CP07 & "以前提出" & m_CPM04 & "。本所代　" & custtype & m_CPM04 & "所需之費用為新台幣　　萬　　仟元整。倘　" & custtype & "續行本案，敬請於　　年　　月　　日以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向該國專利局提出。若逾上述與本所聯絡的期限，除必要之延期費外，本所將另外增收服務費。事關　" & custtype & "權益，尚請特加留意是幸。"
                        g_WordAp.Selection.TypeText "　　五、本案依法應於" & m_CP07 & "以前提出" & m_CPM04 & "。本所代　" & custtype & m_CPM04 & "所需之費用為新台幣　　萬　　仟元整。倘　" & custtype & "續行本案，敬請於" & m_AppDate & "以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向該國專利局提出。若逾上述與本所聯絡的期限，除必要之延期費外，本所將另外增收服務費，另外，本所服務費也會因匯率變動關係而有所調整。事關　" & custtype & "權益，尚請特加留意是幸。"
                     Else
                        'Modified by Morgan 2021/9/3
                        'g_WordAp.Selection.TypeText "　　五、本案依法應於" & m_CP07 & "以前提出" & m_CPM04 & "。本所代　" & custtype & m_CPM04 & "所需之費用為新台幣　　萬　　仟元整。倘　" & custtype & "續行本案，敬請於　　年　　月　　日以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向該國專利局提出。若逾上述與本所聯絡的期限，除必要之延期費外，本所將另外增收服務費。事關　" & custtype & "權益，尚請特加留意是幸。"
                        g_WordAp.Selection.TypeText "　　五、本案依法應於" & m_CP07 & "以前提出" & m_CPM04 & "。本所代　" & custtype & m_CPM04 & "所需之費用為新台幣　　萬　　仟元整。倘　" & custtype & "續行本案，敬請於" & m_AppDate & "以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向該國專利局提出。若逾上述與本所聯絡的期限，除必要之延期費外，本所將另外增收服務費，另外，本所服務費也會因匯率變動關係而有所調整。事關　" & custtype & "權益，尚請特加留意是幸。"
                     End If
                  'end 2015/9/18
                     'add by sonia 2021/4/16 加拿大核駁期限延期很嚴格,所以加一段
                     If pa(9) = "102" Then
                        g_WordAp.Selection.TypeParagraph
                        g_WordAp.Selection.TypeParagraph
                        g_WordAp.Selection.TypeText "*加拿大專利申請案之延期須備具正當理由，一旦專利局認為不符延期規定，本案將被視為放棄，建議盡早決定，以便在期限前及時提出答辯。"
                     End If
                     'end 2021/4/16
                  End If
                  If rsA.State <> adStateClosed Then rsA.Close
                  Set rsA = Nothing
                  '2015/4/16 END
                  '2011/3/16 end
               Case "12"      'CFP最終核駁
                  'Modified by Morgan 2021/9/3 +NP23
                  'StrSQLa = "Select CP07 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1006' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C'"
                  StrSQLa = "Select CP07,NP23 FROM CaseProgress,nextprogress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1006' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' and np01(+)=cp09 and np02(+)=cp01 and np03(+)=cp02 and np04(+)=cp03 and np05(+)=cp04 and np06(+) is null"
                  'end 2021/9/3
                  rsA.CursorLocation = adUseClient
                  rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                  If rsA.RecordCount > 0 Then
                     m_CP07 = Mid(rsA.Fields(0).Value, 1, 4) & "年" & Mid(rsA.Fields(0).Value, 5, 2) & "月" & Mid(rsA.Fields(0).Value, 7, 2) & "日"
                     m_DueDT = ChangeWDateStringToWString(DateAdd("M", -1, ChangeWStringToWDateString(DBDATE(rsA.Fields(0).Value))))
                     m_DueDate = Mid(m_DueDT, 1, 4) & "年" & Mid(m_DueDT, 5, 2) & "月" & Mid(m_DueDT, 7, 2) & "日"
                     'Added by Morgan 2021/9/3 約定期限
                     If rsA("NP23") > 0 Then
                        m_AppDate = Mid(rsA("NP23"), 1, 4) & "年" & Mid(rsA("NP23"), 5, 2) & "月" & Mid(rsA("NP23"), 7, 2) & "日"
                     End If
                     'end 2021/9/3
                  End If
                  If rsA.State <> adStateClosed Then rsA.Close
                  Set rsA = Nothing
                  'Modified by Morgan 2018/1/3 EPC專利局->EPO --玫音
                  'Modified by Morgan 2024/5/16 EPO改半形--郭
                  'Modified by Morgan 2025/9/25 調整內容--柏翰/郭
                  g_WordAp.Selection.TypeText "　　一、　" & custtype & "前委託本所代理向" & nationname & "申請之" & AppNo & "「" & CASENAME & "」" & casetype4 & casetype1 & casetype2 & casetype3 & "案(本所案號" & CaseNo & ")，頃接" & nationname & "專利代理人轉來" & IIf(pa(9) = "221", "EPO", nationname & "專利局") & "所做之最終核駁審定書。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　二、茲隨函附上最終核駁審定書與前案影本乙份，敬請存查。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　三、審查意見："
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　四、本所意見："
                  g_WordAp.Selection.TypeParagraph
                  'Added by Morgan 2015/5/6
                  g_WordAp.Selection.Font.ColorIndex = wdRed
                  'Modified by Morgan 2025/9/25 調整內容--柏翰/郭
                  g_WordAp.Selection.TypeText "　　《註：請工程師在決定建議客戶提出請求繼續審查或是答辯以前，先根據附錄第一點與第二點，確認本案下一程序是否涉及請求項、說明書或圖式的實質修正，若涉及實質修正，則應建議客戶提出請求繼續審查，而非答辯；若建議答辯，請在本節建議客戶特別注意須在最終審定發出後三個月以內提出答辯；請閱畢後刪除本段註記。》"
                  g_WordAp.Selection.Font.ColorIndex = wdAuto
                  'end 2015/5/6
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　關於美國最終審定後之相關程序，詳請參看附錄。"
                  g_WordAp.Selection.TypeParagraph
                  'Modified by Morgan 2021/9/3
                  'g_WordAp.Selection.TypeText "　　五、本案依法應於" & m_CP07 & "以前提出請求繼續審查或答辯，若　" & custtype & "期望能獲得審查委員之建議性處分，且能在" & m_DueDate & "以前提出答辯，則有利於在本案的後續程序。本所代　" & custtype & "提出繼續審查所需之費用為新台幣　　萬　　仟元整，另本所代　" & custtype & "答辯所需之費用為新台幣　　萬　　仟元整。　" & custtype & "如欲進行後續程序，為避免衍生延期費，敬請於　　年　　月　　日以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向該國專利局提出，若逾上述與本所聯絡的期限，除必要之延期費外，本所將另外增收服務費。此外，如逾期未蒙見覆，本所即認為　" & custtype & "不擬續行此案，事關　" & custtype & "權益，尚請特加留意是幸。"
                  'Modified by Morgan 2025/9/25 調整內容--柏翰/郭
                  'g_WordAp.Selection.TypeText "　　五、本案依法應於" & m_CP07 & "以前提出請求繼續審查或答辯，若　" & custtype & "期望能獲得審查委員之建議性處分，且能在" & m_DueDate & "以前提出答辯，則有利於在本案的後續程序。本所代　" & custtype & "提出繼續審查所需之費用為新台幣　　萬　　仟元整，另本所代　" & custtype & "答辯所需之費用為新台幣　　萬　　仟元整。　" & custtype & "如欲進行後續程序，為避免衍生延期費，敬請於" & m_AppDate & "以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向該國專利局提出，若逾上述與本所聯絡的期限，除必要之延期費外，本所將另外增收服務費，另外，本所服務費也會因匯率變動關係而有所調整。此外，如逾期未蒙見覆，本所即認為　" & custtype & "不擬續行此案，事關　" & custtype & "權益，尚請特加留意是幸。"
                  g_WordAp.Selection.TypeText "　　五、本案依法應於" & m_CP07 & "以前提出請求繼續審查或答辯，如果　" & custtype & "期望能透過提出答辯獲得審查委員之建議性處分，則該答辯若能在" & m_DueDate & "以前提出，將有利於在本案的後續程序。本所代　" & custtype & "提出繼續審查所需之費用為新台幣　　萬　　仟元整，本所代　" & custtype & "答辯所需之費用為新台幣　　萬　　仟元整。　" & custtype & "如欲進行後續程序，為避免衍生延期費，敬請於" & m_AppDate & "以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向該國專利局提出，若逾上述與本所聯絡的期限，除必要之延期費外，本所將另外增收服務費，另外，本所服務費也會因匯率變動關係而有所調整。此外，如逾期未蒙見覆，本所即認為　" & custtype & "不擬續行此案，事關　" & custtype & "權益，尚請特加留意是幸。"
                  '2010/1/6 cancel by sonia
                  'g_WordAp.Selection.TypeParagraph
                  'g_WordAp.Selection.TypeText "　　六、此外，若　" & custtype & "欲提出繼續審查，本所代　" & custtype & "提出繼續審查所需之費用為新台幣　　萬　　仟元整，在此一併向　" & custtype & "報告。"
                  '2010/1/6 end
               Case "13"      'CFP通知要求選取
                  'Modified by Morgan 2021/9/3 +NP23
                  'Modified by Morgan 2021/9/3
                  'StrSQLa = "Select CP07 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1206' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C'"
                  StrSQLa = "Select CP07,NP23 FROM CaseProgress,nextprogress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1206' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' and np01(+)=cp09 and np02(+)=cp01 and np03(+)=cp02 and np04(+)=cp03 and np05(+)=cp04 and np06(+) is null"
                  'end 2021/9/3
                  rsA.CursorLocation = adUseClient
                  rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                  If rsA.RecordCount > 0 Then
                     m_CP07 = Mid(rsA.Fields(0).Value, 1, 4) & "年" & Mid(rsA.Fields(0).Value, 5, 2) & "月" & Mid(rsA.Fields(0).Value, 7, 2) & "日"
                    'Added by Morgan 2021/9/3 約定期限
                     If rsA("NP23") > 0 Then
                        m_AppDate = Mid(rsA("NP23"), 1, 4) & "年" & Mid(rsA("NP23"), 5, 2) & "月" & Mid(rsA("NP23"), 7, 2) & "日"
                     End If
                     'end 2021/9/3
                  End If
                  If rsA.State <> adStateClosed Then rsA.Close
                  Set rsA = Nothing
                  'Modify By Sindy 2012/8/29 修改定稿內容
                  g_WordAp.Selection.TypeText "　　一、　" & custtype & "前委託本所代理向" & nationname & "申請之" & AppNo & "「" & CASENAME & "」" & casetype4 & casetype1 & casetype2 & casetype3 & "案(本所案號" & CaseNo & ")，頃接該國專利代理人轉來該國專利局之選取通知，謂本案包含有一種以上之發明。"
                  g_WordAp.Selection.TypeParagraph
                  'Modified by Morgan 2024/9/4 --柏翰
                  'g_WordAp.Selection.TypeText "　　二、茲隨函附上審定書乙份，敬請存查。"
                  g_WordAp.Selection.TypeText "　　二、茲隨函附上選取通知乙份，敬請存查。"
                  'end 2024/9/4
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　三、審查意見："
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　四、本所意見："
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　若　" & custtype & "未能於期限前回覆選取通知，則本案將被視為放棄，關於美國選取之相關程序及規定，詳請參看附錄。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  '2009/10/26 modify by sonia 取消分割案報價陸萬陸仟元由工程師自行填入CFP-018981
                  'Modified by Morgan 2020/2/11 若有主張國際優先權時也要帶費用 Ex:CFP-029749 --郭
                  strExc(1) = ""
                  'Modified by Morgan 2020/2/20 +IDS報價--郭
                  If PUB_ChkCPExist(pa(), "106", 2) = True Then
                     'strExc(1) = "，主張優先權之費用為XX元整"
                     strExc(1) = "，主張優先權－新台幣　　　元整"
                  End If
                  'g_WordAp.Selection.TypeText "　　五、本案依據該國法律規定，應於" & m_CP07 & "以前提出選取。本所代　" & custtype & "提出選取所需之費用為新台幣　　萬　　仟元整。　" & custtype & "如欲提出選取，敬請於　　年　　月　　日以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向該國專利局提出。若逾上述與本所聯絡的期限，除必要之延期費外，本所將另外增收服務費。另外，本所代　" & custtype & "針對每一組未選取之申請專利範圍分別提出分割案所需之費用為新台幣　　萬　　仟元整" & strExc(1) & "，在此一併向　" & custtype & "報告。事關　" & custtype & "權益，尚請特加留意是幸。"
                  'Modified by Morgan 2021/9/3
                  'g_WordAp.Selection.TypeText "　　五、本案依據該國法律規定，應於" & m_CP07 & "以前提出選取。本所代　" & custtype & "提出選取所需之費用為新台幣　　萬　　仟元整。　" & custtype & "如欲提出選取，敬請於　　年　　月　　日以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向該國專利局提出。若逾上述與本所聯絡的期限，除必要之延期費外，本所將另外增收服務費。另外，本所代　" & custtype & "針對每一組未選取之申請專利範圍分別提出分割案所需之費用：分割申請－新台幣　　萬　　仟元整，IDS－新台幣　　　元整" & strExc(1) & "，在此一併向　" & custtype & "報告。事關　" & custtype & "權益，尚請特加留意是幸。"
                  'Modified by Morgan 2022/10/19
                  'g_WordAp.Selection.TypeText "　　五、本案依據該國法律規定，應於" & m_CP07 & "以前提出選取。本所代　" & custtype & "提出選取所需之費用為新台幣　　萬　　仟元整。　" & custtype & "如欲提出選取，敬請於" & m_AppDate & "以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向該國專利局提出。若逾上述與本所聯絡的期限，除必要之延期費外，本所將另外增收服務費。另外，本所代　" & custtype & "針對每一組未選取之申請專利範圍分別提出分割案所需之費用：分割申請－新台幣　　萬　　仟元整，IDS－新台幣　　　元整" & strExc(1) & "，在此一併向　" & custtype & "報告。事關　" & custtype & "權益，尚請特加留意是幸。"
                  g_WordAp.Selection.TypeText "　　五、本案依據該國法律規定，應於" & m_CP07 & "以前提出選取。本所代　" & custtype & "提出選取所需之費用為新台幣　　萬　　仟元整。　" & custtype & "如欲提出選取，敬請於" & m_AppDate & "以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向該國專利局提出。若逾上述與本所聯絡的期限，除必要之延期費外，本所將另外增收服務費。另外，本所代　" & custtype & "針對每一組未選取之申請專利範圍分別提出分割案所需之費用：分割申請－新台幣　　萬　　仟元整" & strExc(1) & "，IDS的費用依IDS提出的時點另外報價，在此一併向　" & custtype & "報告。事關　" & custtype & "權益，尚請特加留意是幸。"
                  'end 2022/10/19
                  'end 2020/2/11
               Case "14"      '檢索報告
                  
                  '2009/3/3 add by sonia 加PCT案
                  'Modify by Morgan 2010/7/19 +CFP的PCT案
                  'If pa(1) = "P" And pa(9) = "056" Then
                  If (pa(1) = "P" Or pa(1) = "CFP") And pa(9) = "056" Then
                     '抓最早優先權日
                     strFirstPriDate = PUB_GetFirstPriDate(pa)
                     '修正法定期限=申請日(最早優先權日)+16個月或檢索報告來函收文日+2個月取較晚者
                     If strFirstPriDate <> "" Then
                        m_CP07 = CompDate(1, 16, TransDate(strFirstPriDate, 2))
                     Else
                        m_CP07 = CompDate(1, 16, TransDate(pa(10), 2))
                     End If
                     
                     '抓檢索報告來函收文日+2個月
                     'Add by Morgan 2010/7/19
                     If pa(1) = "CFP" Then
                        StrSQLa = "Select CP64,CP133 FROM CaseProgress WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "'" & _
                           " AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP10='1209' AND CP27 IS NULL "
                     Else
                     'end 2010/7/19
                     
                        'StrSQLa = "Select C2.CP05 FROM CaseProgress C1,CaseProgress C2 WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                        '          " AND C1.CP10='901' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP43=C2.CP09(+) AND '1209'=C2.CP10(+) "
                        '2009/8/6 MODIFY BY SONIA 改抓檢索報告之進度備註的機關發文日(等請作單)
                        StrSQLa = "Select C2.CP64,C2.CP133 FROM CaseProgress C1,CaseProgress C2 WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                                  " AND C1.CP10='901' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP43=C2.CP09(+) AND '1209'=C2.CP10(+) "
                        '2009/8/6 END
                        'Added by Morgan 2018/9/3 已改C類不發文，不新增告代 Ex:P-119984 --茹曣
                        StrSQLa = StrSQLa & " union Select CP64,CP133 FROM CaseProgress WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "'" & _
                           " AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP10='1209' AND CP27 IS NULL "
                        'end 2018/9/3
                     End If
                     
                     rsA.CursorLocation = adUseClient
                     rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                     If rsA.RecordCount > 0 Then
                        'Add by Morgan 2010/7/19
                        If Not IsNull(rsA.Fields(1)) Then
                           m_DueDT = CompDate(1, 2, rsA.Fields(1))
                        Else
                        'end 2010/7/19
                           If rsA.Fields(0).Value <> "" Then
                              'm_DueDT = CompDate(1, 2, TransDate(rsA.Fields(0).Value, 2))
                              '2009/8/6 MODIFY BY SONIA 改抓檢索報告之進度備註的機關發文日
                              m_DueDT = CompDate(1, 2, TransDate(Val(Mid(rsA.Fields(0).Value, 7, 8)), 2))
                              '2009/8/6 END
                           End If
                        End If
                     End If
                     If rsA.State <> adStateClosed Then rsA.Close
                     Set rsA = Nothing
                     '修正法定期限=申請日(最早優先權日)+16個月或檢索報告來函收文日+2個月取較晚者
                     If m_DueDT > m_CP07 Then
                        m_CP07 = m_DueDT
                     End If
                     m_CP07 = CompDate(2, -10, TransDate(m_CP07, 2))     '本所期限=法定期限-10天,再取工作天
                     m_CP07 = PUB_GetWorkDay1(m_CP07, True)
                     m_DueDate = CompDate(2, -7, TransDate(m_CP07, 2))   '約定期限=本所期限-2天
                     m_DueDate = PUB_GetWorkDay1(m_DueDate, True)        'add by sonia 2020/12/24 取工作天
                     
                     m_CP07 = Mid(m_CP07, 1, 4) & "年" & Mid(m_CP07, 5, 2) & "月" & Mid(m_CP07, 7, 2) & "日"
                     m_DueDate = Mid(m_DueDate, 1, 4) & "年" & Mid(m_DueDate, 5, 2) & "月" & Mid(m_DueDate, 7, 2) & "日"
                     'Modified by Morgan 2012/1/3 調整段落空白行--游登銘
                     g_WordAp.Selection.TypeText "　　一、　" & custtype & "前委託本所辦理之" & AppNo & "「" & CASENAME & "」" & nationname & casetype4 & casetype1 & casetype2 & casetype3 & "案(本所案號" & CaseNo & ")，頃接獲代理人轉來本案之國際檢索報告及國際檢索單位書面意見通知書。"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　二、茲隨函附上國際檢索報告、國際檢索單位書面意見通知書與前案影本乙份，敬請存查。"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　三、審查意見："
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　再者，審查員修改本案摘要如檢索報告第2頁所示，針對修改後的摘要若有不同意見可在國際檢索報告發文日起一個月內向檢索局提出，惟由於摘要內容並不會影響本案的實質權利範圍，一般可接受審查員的修改內容。"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　(註：前段內容請依實際檢索報告內容予以保留或刪除)"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　四、本所意見："
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　(註：有部份項次不具專利要件)"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeParagraph
                     'Modified by Morgan 2015/9/14 "本案依法可於"->"本案可於" --郭
                     'Modified by Morgan 2020/8/6 --郭
                     'g_WordAp.Selection.TypeText "　　由於該國際檢索報告中指出本案之申請專利範圍第  至  項不具新穎性、第  至  項不具創造性，　" & custtype & "可根據目前國際檢索單位的書面審查意見，依據PCT條約第19條提出回覆意見並一併修正本案申請專利範圍，則本案之國際公開階段將一併公開該等修正本，依據PCT條約第19條之修正本的期限為傳送國際檢索報告之日起兩個月內或本案最早優先權日起十六個月內(以較晚屆滿期限為準)，故本案可於" & m_CP07 & "前遞交修正本，倘若　" & custtype & "欲辦理國際階段修正程序，敬請於" & m_DueDate & "以前與本所聯絡。"
                     g_WordAp.Selection.TypeText "　　由於該國際檢索報告中指出本案之申請專利範圍第  至  項不具新穎性、第  至  項不具創造性，　" & custtype & "可根據目前國際檢索單位的書面審查意見，依據PCT條約第19條提出修正申請專利範圍，則本案之國際公開階段將一併公開該等修正本，依據PCT條約第19條之修正期限為傳送國際檢索報告之日起兩個月內或本案最早優先權日起十六個月內(以較晚屆滿期限為準)，故本案可於" & m_CP07 & "前遞交修正本，倘若　" & custtype & "欲辦理國際階段修正程序，敬請於" & m_DueDate & "以前與本所聯絡。"
                     'end 2020/8/6
                     g_WordAp.Selection.TypeParagraph
                     'Modified by Morgan 2020/8/6 --郭
                     'g_WordAp.Selection.TypeText "　　另，本案可選擇暫不修正，於進入各國國家階段後再依據各國的審定書作出回覆，而屆時各國審查委員將依其主觀判斷選擇是否引用國際檢索單位的審查意見。"
                     g_WordAp.Selection.TypeText "　　另，本案可選擇暫不修正，於進入各國國家階段後再依據各國的審查意見作出回覆，而屆時各國審查委員將依其主觀判斷選擇是否引用國際檢索單位的審查意見。"
                     'end 2020/8/6
                     g_WordAp.Selection.TypeParagraph
                     'Modified by Morgan 2018/9/26 -- 郭雅娟
                     'g_WordAp.Selection.TypeText "　　再者，由於本案目前尚未提出國際初步審查請求，國際初步審查請求為非強制性，　" & custtype & "可自由選擇是否要提申，倘若欲提出國際初步審查請求，則可根據目前國際檢索報告的審查意見，於日後提出國際初步審查請求時一併修正本案，如此可讓審查委員依新修正的內容進行國際初步審查，使本案有機會先一步避開國際檢索報告中所引用的前案。"
                     'Modified by Morgan 2020/8/6 --郭
                     'g_WordAp.Selection.TypeText "　　再者，國際初步審查請求為非強制性，　" & custtype & "可自由選擇是否要提申，倘若欲提出國際初步審查請求，則可根據目前國際檢索報告的審查意見，於日後提出國際初步審查請求時一併修正本案，如此可讓審查委員依新修正的內容進行國際初步審查，使本案有機會先一步避開國際檢索報告中所引用的前案。"
                     g_WordAp.Selection.TypeText "　　再者，國際初步審查請求為非強制性，　" & custtype & "可自由選擇是否要提申，倘若欲提出國際初步審查請求，則可根據目前國際檢索報告的審查意見，於日後提出國際初步審查請求時一併提交答覆意見及修正本案申請內容，如此可讓審查委員依新修正的內容進行國際初步審查，使本案有機會先一步避開國際檢索報告中所引用的前案。"
                     'end 2020/8/6
                     'end 2018/9/26
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "===================================================================="
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　四、本所意見："
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　(註：全部項次均具專利要件)"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeParagraph
                     'Modified by Morgan 2015/9/14 "本案依法可於"->"本案可於" --郭
                     g_WordAp.Selection.TypeText "　　依據PCT條約第19條的規定，於傳送國際檢索報告之日起兩個月內或本案最早優先權日起十六個月內(以較晚屆滿期限為準)可提出國際階段之申請專利範圍修正本，故本案可於" & m_CP07 & "前遞交修正本，倘若　" & custtype & "基於任何考量欲辦理國際階段修正程序，敬請於" & m_DueDate & "以前與本所聯絡。"
                     g_WordAp.Selection.TypeParagraph
                     'Removed by Morgan 2018/9/26 --郭
                     'g_WordAp.Selection.TypeText "　　再者，由於本案目前尚未提出國際初步審查請求，國際初步審查請求為非強制性，　" & custtype & "可自由選擇是否要提出請求，關於是否請求國際初步審查所可能產生的影響，請參閱次段所示。"
                     'g_WordAp.Selection.TypeParagraph
                     'end 2018/9/26
                     'add by sonia 2016/12/30 游經理
                     g_WordAp.Selection.TypeText "===================================================================="
                      g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "〔註：前兩段的「四、本所意見」內容請依案情選擇其一，完成後整理各段落並刪除雙虛線。〕"
                     g_WordAp.Selection.TypeParagraph
                    'end 2016/12/30
                     g_WordAp.Selection.TypeParagraph
                     '進入4個成員國法定期限=申請日(最早優先權日)+20個月
                     If strFirstPriDate <> "" Then
                        m_CP07 = CompDate(1, 20, TransDate(strFirstPriDate, 2))
                     Else
                        m_CP07 = CompDate(1, 20, TransDate(pa(10), 2))
                     End If
                     m_CP07 = Mid(m_CP07, 1, 4) & "年" & Mid(m_CP07, 5, 2) & "月" & Mid(m_CP07, 7, 2) & "日"
                     'Modified by Morgan 2015/1/28 第二個二十個月改為三十個月 -- 郭雅娟
                     'Modified by Morgan 2018/9/26 --郭
                     'g_WordAp.Selection.TypeText "　　五、因目前仍然有3個PCT成員國(盧森堡Luxembourg、坦尚尼亞Tanzania、烏干達Uganda)保留並執行PCT條約第2章的規定，即申請人倘若未提出國際初步審查請求，則必須於國際申請日(如有優先權日者，以優先權日為準)二十個月內完成進入上述3國的國家階段手續，倘若申請人有提出國際初步審查請求，則可於國際申請日(如有優先權日者，以優先權日為準)三十個月內完成進入上述3國的國家階段手續。"
                     g_WordAp.Selection.TypeText "　　五、因目前仍然有2個PCT成員國(盧森堡Luxembourg、坦尚尼亞Tanzania)保留並執行PCT條約第2章的規定，即申請人倘若未提出國際初步審查請求，則必須於國際申請日(如有優先權日者，以優先權日為準)二十個月內完成進入上述2國的國家階段手續，倘若申請人有提出國際初步審查請求，則可於國際申請日(如有優先權日者，以優先權日為準)三十個月內完成進入上述2國的國家階段手續。"
                     'end 2018/9/26
                     'end 2015/1/28
                     g_WordAp.Selection.TypeParagraph
                     'Modified by Morgan 2018/9/26 --郭
                     'g_WordAp.Selection.TypeText "　　因此，倘若　" & custtype & "認為本案將來有機會進入上述3國，則應慎重考慮提出國際初步審查請求，或應考慮提早於" & m_CP07 & "前完成進入國家階段手續，事關　" & custtype & "權益，尚祈能特加留意是幸。"
                     g_WordAp.Selection.TypeText "　　因此，倘若　" & custtype & "認為本案將來有機會進入上述2國，則應慎重考慮提出國際初步審查請求，或應考慮提早於" & m_CP07 & "前完成進入國家階段手續，事關　" & custtype & "權益，尚祈能特加留意是幸。"
                     'end 2018/9/26
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　六、以上分析意見僅供參考，如對前述內容有任何疑問，請隨時洽詢本所，本所將竭盡全力為　" & custtype & "服務。"
                     'Added by Morgan 2013/4/18
                     strExc(1) = "": strExc(2) = ""
                     If Cls001GetNextProgressDate(pa(1), pa(2), pa(3), pa(4), "119", strExc(1), strExc(2)) = True Then
                        'modify by sonia 2020/2/20 TranslateKeyWord(incCNV_CHINESE_MINKO...加傳本所案號,以判斷日期欄之民國或西元格式
                        strExc(1) = TranslateKeyWord(incCNV_CHINESE_MINKO, strExc(1), "", pa(1) & pa(2) & pa(3) & pa(4))
                        strExc(2) = TranslateKeyWord(incCNV_CHINESE_MINKO, strExc(2), "", pa(1) & pa(2) & pa(3) & pa(4))
                     End If
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　七、本案進入各國國家階段之最後期限為" & strExc(2) & "，若欲辦理，煩請儘早於" & strExc(1) & "前通知本所。"
                     'end 2013/4/18
                     
                  '2009/3/3 END
                  ElseIf pa(1) = "CFP" Then
                     'Modified by Morgan 2021/9/3 +NP23
                     'StrSQLa = "Select CP07 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                               " AND CP10='1209' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C'"
                     StrSQLa = "Select CP07,NP23 FROM CaseProgress,NEXTPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                               " AND CP10='1209' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' and np01(+)=cp09 and np02(+)=cp01 and np03(+)=cp02 and np04(+)=cp03 and np05(+)=cp04 and np06(+) is null "
                     'end 2021/9/3
                     rsA.CursorLocation = adUseClient
                     rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                     If rsA.RecordCount > 0 Then
                        If rsA.Fields(0).Value <> "" Then
                           m_CP07 = Mid(rsA.Fields(0).Value, 1, 4) & "年" & Mid(rsA.Fields(0).Value, 5, 2) & "月" & Mid(rsA.Fields(0).Value, 7, 2) & "日"
                        Else
                           m_CP07 = "　　年　　月　　日"
                        End If
                        'Added by Morgan 2021/9/3 約定期限
                        If rsA("NP23") > 0 Then
                           m_AppDate = Mid(rsA("NP23"), 1, 4) & "年" & Mid(rsA("NP23"), 5, 2) & "月" & Mid(rsA("NP23"), 7, 2) & "日"
                        End If
                        'end 2021/9/3
                     End If
                     If rsA.State <> adStateClosed Then rsA.Close
                     Set rsA = Nothing
                     'Modified by Morgan 2018/1/3 EPC專利局->EPO --玫音
                     'Modified by Morgan 2024/5/16 EPC案調整內容,其他國家不變--郭
                     'g_WordAp.Selection.TypeText "　　一、　" & custtype & "前委託本所代理向" & nationname & "申請之" & AppNo & "「" & CASENAME & "」" & casetype4 & casetype1 & casetype2 & casetype3 & "案(本所案號" & CaseNo & ")，頃接" & nationname & "專利代理人轉來" & IIf(nationname = "ＥＰＣ", "ＥＰＯ", nationname & "專利局") & "之新穎性調查報告。"
                     'g_WordAp.Selection.TypeParagraph
                     'g_WordAp.Selection.TypeText "　　二、茲隨函附上審定書與前案影本乙份，敬請存查。"
                     If pa(9) = "221" Then
                        g_WordAp.Selection.TypeText "　　一、　" & custtype & "前委託本所代理向" & nationname & "申請之" & AppNo & "「" & CASENAME & "」" & casetype4 & casetype1 & casetype2 & casetype3 & "案(本所案號" & CaseNo & ")，頃接" & nationname & "專利代理人轉來EPO之擴大檢索報告。"
                        g_WordAp.Selection.TypeParagraph
                        g_WordAp.Selection.TypeText "　　二、茲隨函附上擴大檢索報告與前案影本乙份，敬請存查。"
                     Else
                        g_WordAp.Selection.TypeText "　　一、　" & custtype & "前委託本所代理向" & nationname & "申請之" & AppNo & "「" & CASENAME & "」" & casetype4 & casetype1 & casetype2 & casetype3 & "案(本所案號" & CaseNo & ")，頃接" & nationname & "專利代理人轉來" & nationname & "專利局之新穎性調查報告。"
                        g_WordAp.Selection.TypeParagraph
                        g_WordAp.Selection.TypeText "　　二、茲隨函附上審定書與前案影本乙份，敬請存查。"
                     End If
                     'end 2024/5/16
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　三、審查意見："
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　四、本所意見："
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeParagraph
                     '2015/4/14 MODIFY BY SONIA 依是否有美國關聯案決定要不要有 前案揭露聲明 的段落 CFP-027079, 且美國案依大小微個體之分,費用不同
                     'g_WordAp.Selection.TypeText "　　五、"
                     'g_WordAp.Selection.TypeParagraph
                     'g_WordAp.Selection.TypeParagraph
                     ''modify by sonia 2013/6/4 原為壹萬玖仟元
                     'g_WordAp.Selection.TypeText "　　按依照美國專利法之規定，申請人有義務將所知與本案有關的前案主動提交美國專利局審查員作為審查本案之參考，若有明知而未陳報而被發現或者遭他人舉報者，將造成對本案不利之後果，即美國專利法所規定之前案揭露聲明（INFORMATION DISCLOSURE STATEMENT, IDS）。經查　" & custtype & "另申請有與本案相關之美國專利申請案(本所案號CFP-　　)，因此依美國專利法之規定，應於三個月內主動將本次檢索報告向美國專利局提出，以符合美國專利法之規定，避免對相關之美國專利案產生不利的後果。另外美國提出前案揭露所需費用會依所提出時間不同而有所不同，若在該美國專利案第一次審定書發出前提出前案揭露，本所代　" & custtype & "提出前案揭露的費用為壹萬柒仟元整（經查該與本案相關之美國專利案目前並未收到第一次審定書），而若在接獲第一次審定書後方才提出前案揭露，則需費用貳萬陸仟元整，在此一併告知　" & custtype & "，事關　" & custtype & "權益，尚請特加留意是幸"
                     'g_WordAp.Selection.TypeParagraph
                     'g_WordAp.Selection.TypeParagraph
                     'g_WordAp.Selection.TypeText "　　六、" & custtype & "若現今即願意修正本案，本所代　" & custtype & "提出修正之費用為新台幣　　萬　　仟元整，實體審查費用為新台幣　　萬　　仟元整。倘　" & custtype & "欲續行本案，則本案依法可於" & m_CP07 & "前向" & nationname & "專利局作出回覆，故敬請於　　年　　月　　日以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向專利局提出。若逾上述與本所聯絡的期限，本所之服務費將會因匯率變動關係而有所調整。事關　" & custtype & "權益，尚請特加留意是幸。"
                     'Modified by Morgan 2018/5/30
                     'StrSQLa = "SELECT CR05||'-'||CR06||DECODE(CR07,'0',NULL,'-'||CR07||'-'||CR08) CR101NO,DECODE(SUBSTR(PA91,INSTR(PA91,'個體')-1,1),'大','貳萬陸仟元','小','貳萬參仟元','微','貳萬壹仟元','') FEE FROM CASERELATION,PATENT" & _
                               " WHERE CR01='" & pa(1) & "' AND CR02='" & pa(2) & "' AND CR03='" & pa(3) & "' AND CR04='" & pa(4) & "'" & _
                               " AND CR05=PA01(+) AND CR06=PA02(+) AND CR07=PA03(+) AND CR08=PA04(+) AND PA09='101' "
                     ''2013/4/26 end
                     'rsA.CursorLocation = adUseClient
                     'rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                     ''有美國關聯案
                     'If rsA.RecordCount > 0 Then
                     If fnUsIdsChk(rsA) = True Then
                     'end 2018/5/30
                        g_WordAp.Selection.Font.ColorIndex = wdRed
                        g_WordAp.Selection.TypeText "　　五、前案揭露聲明（INFORMATION DISCLOSURE STATEMENT, IDS）《註：請承辦工程師確認本節IDS內容是否留存。若與本案相關之美國專利案已取得專利證書，或本案的相關前案均已見於相關美國專利案，則本節IDS內容無須留存。請工程師閱畢後刪除本段註記。》"
                        g_WordAp.Selection.TypeParagraph
                        'Modified by Morgan 2016/11/9 不要帶金額 --郭
                        'g_WordAp.Selection.TypeText "　　按依照美國專利法之規定，申請人有義務將所知與本案有關的前案主動提交美國專利局審查員作為審查與本案相關之美國專利案之參考，即美國專利法所規定之前案揭露聲明。若有明知而未陳報而被發現或者遭他人舉報者，將對相關美國專利案造成不利之後果。經查　" & custtype & "另申請有與本案相關之美國專利申請案（本所案號" & rsA.Fields("CR101NO") & "），因此依美國專利法之規定，應於三個月內主動將本次審查意見通知書（檢索報告）所附前案向美國專利局提出，以符合美國專利法之規定，避免對相關美國專利案產生不利的後果。另外美國提出前案揭露所需費用會依所提出時間不同而有所不同，若在該美國專利案第一次審定書發出前提出前案揭露聲明，本所代　" & custtype & "提出前案揭露聲明的費用為壹萬柒仟元整（經查該與本案相關之美國專利案目前並未收到第一次審定書），而若在發出第一次審定書後方才提出前案揭露聲明，則需費用" & rsA.Fields("FEE") & "整，在此一併告知　" & custtype & "，事關　" & custtype & "權益，尚請特加留意是幸。"
                        'Modified by Morgan 2019/12/23 內容調整--郭
                        'g_WordAp.Selection.TypeText "　　按依照美國專利法之規定，申請人有義務將所知與本案有關的前案主動提交美國專利局審查員作為審查與本案相關之美國專利案之參考，即美國專利法所規定之前案揭露聲明。若有明知而未陳報而被發現或者遭他人舉報者，將對相關美國專利案造成不利之後果。經查　" & custtype & "另申請有與本案相關之美國專利申請案（本所案號" & rsA.Fields("CR101NO") & "），因此依美國專利法之規定，應於三個月內主動將本次審查意見通知書（檢索報告）所附前案向美國專利局提出，以符合美國專利法之規定，避免對相關美國專利案產生不利的後果。另外美國提出前案揭露所需費用會依所提出時間不同而有所不同，若在該美國專利案第一次審定書發出前提出前案揭露聲明，本所代　" & custtype & "提出前案揭露聲明的費用為新台幣　　萬　　仟元整（經查該與本案相關之美國專利案目前並未收到第一次審定書），而若在發出第一次審定書後方才提出前案揭露聲明，則需費用新台幣　　萬　　仟元整，在此一併告知　" & custtype & "，事關　" & custtype & "權益，尚請特加留意是幸。"
                        'Modified by Morgan 2020/8/6 內容調整--郭
                        'g_WordAp.Selection.TypeText "　　按依照美國專利法之規定，申請人有義務將所知與本案有關的前案主動提交美國專利局審查員作為審查與本案相關之美國專利案之參考，即美國專利法所規定之前案揭露聲明。若有明知而未陳報而被發現或者遭他人舉報者，將對相關美國專利案造成不利之後果。經查　" & custtype & "另申請有與本案相關之美國專利申請案（本所案號" & rsA.Fields("CR101NO") & "），因此依美國專利法之規定，應於首次得知的三個月內主動將本次審查意見通知書（檢索報告）所附前案向美國專利局提出，以符合美國專利法之規定，避免對相關美國專利案產生不利的後果，此階段所需費用為新台幣　　萬　　仟元整；惟若是超過三個月後且美國專利申請案已發出審查意見通知書才提出前案揭露聲明者，此階段所需費用為新台幣　　萬　　仟元整，在此一併告知，事關　" & custtype & "權益，尚請特加留意是幸。"
                        'Modified by Morgan 2023/10/30 內容調整--郭
                        'g_WordAp.Selection.TypeText "　　按依照美國專利法之規定，申請人有義務將所知與本案有關的前案主動提交美國專利局審查員作為審查與本案相關之美國專利案之參考，即美國專利法所規定之前案揭露聲明。若有明知而未陳報而被發現或者遭他人舉報者，將對相關美國專利案造成不利之後果。經查　" & custtype & "另申請有與本案相關之美國專利申請案（本所案號" & rsA.Fields("CR101NO") & "），因此依美國專利法之規定，應於首次得知的三個月內主動將本次審查意見通知書（檢索報告）所附前案向美國專利局提出，以符合美國專利法之規定，避免對相關美國專利案產生不利的後果，此階段所需費用為新台幣　　萬　　仟元整；惟若是超過三個月後且美國專利申請案已發出審查意見通知書才提出前案揭露聲明者，或是美國專利申請案已接獲最終核駁或核准通知書(在未繳納領證費前)才提出前案揭露聲明者，此階段所需費用為新台幣　　萬　　仟元整，在此一併告知，事關　" & custtype & "權益，尚請特加留意是幸。"
                        'Modified by Morgan 2025/5/22 內容調整--郭
                        'g_WordAp.Selection.TypeText "　　按依照美國專利法之規定，申請人有義務將所知與本案有關的前案主動提交美國專利局審查員作為審查與本案相關之美國專利案之參考，即美國專利法所規定之前案揭露聲明。若有明知而未陳報而被發現或者遭他人舉報者，將對相關美國專利案造成不利之後果。經查　" & custtype & "另申請有與本案相關之美國專利申請案（本所案號" & rsA.Fields("CR101NO") & "），因此依美國專利法之規定，應於首次得知的三個月內主動將本次審查意見通知書（檢索報告）所附前案向美國專利局提出，以符合美國專利法之規定，避免對相關美國專利案產生不利的後果，此階段所需費用為新台幣　　萬　　仟元整(前案以25件為限，每超出5件須加收新台幣2,500元)；惟若是超過三個月後且美國專利申請案已發出審查意見通知書才提出前案揭露聲明者，或是美國專利申請案已接獲最終核駁或核准通知書(在未繳納領證費前)才提出前案揭露聲明者，此階段所需費用為新台幣　　萬　　仟元整(前案以25件為限，每超出5件須加收新台幣2,500元)，在此一併告知，事關　" & custtype & "權益，尚請特加留意是幸。"
                        g_WordAp.Selection.TypeText "　　按依照美國專利法之規定，申請人有義務將所知與本案有關的前案主動提交美國專利局審查員作為審查與本案相關之美國專利案之參考，即美國專利法所規定之前案揭露聲明。若有明知而未陳報而被發現或者遭他人舉報者，將對相關美國專利案造成不利之後果。經查　" & custtype & "另申請有與本案相關之美國專利申請案（本所案號" & rsA.Fields("CR101NO") & "），因此依美國專利法之規定，應於首次得知的三個月內主動將本次審查意見通知書（檢索報告）所附前案向美國專利局提出，以符合美國專利法之規定，避免對相關美國專利案產生不利的後果，此階段所需費用為新台幣　　萬　　仟元整(前案以25件為限，每超出5件須加收新台幣2,500元)；"
                        g_WordAp.Selection.TypeParagraph
                        g_WordAp.Selection.TypeText "　　惟若是超過三個月後且美國專利申請案已發出審查意見通知書才提出前案揭露聲明者，或是雖未超過三個月但美國專利申請案已接獲最終核駁或核准通知書(在未繳納領證費前)才提出前案揭露聲明者，此階段所需費用為新台幣　　萬　　仟元整(前案以25件為限，每超出5件須加收新台幣2,500元);"
                        g_WordAp.Selection.TypeParagraph
                        g_WordAp.Selection.TypeText "　　惟若是超過三個月後且美國專利申請案已接獲最終核駁或核准通知書(在未繳納領證費前)才提出前案揭露聲明者，此階段的程序較為複雜，將依個案情形另外提供報價；"
                        g_WordAp.Selection.TypeParagraph
                        g_WordAp.Selection.TypeText "　　在此一併告知，事關　" & custtype & "權益，尚請特加留意是幸。"
                        'end 2025/5/22
                        'end 2023/10/30
                        'end 2020/8/6
                        g_WordAp.Selection.Font.ColorIndex = wdAuto
                        g_WordAp.Selection.TypeParagraph
                        g_WordAp.Selection.TypeParagraph
                        'Modified by Morgan 2015/9/14 "本案依法可於"->"本案可於" --郭
                        'Modified by Morgan 2018/1/3 EPC專利局->EPO --玫音
                        'Modified by Morgan 2021/9/3
                        'g_WordAp.Selection.TypeText "　　六、" & custtype & "若現今即願意修正本案，本所代　" & custtype & "提出修正之費用為新台幣　　萬　　仟元整，實體審查費用為新台幣　　萬　　仟元整。倘　" & custtype & "欲續行本案，則本案可於" & m_CP07 & "前向" & IIf(nationname = "ＥＰＣ", "ＥＰＯ", nationname & "專利局") & "作出回覆，故敬請於　　年　　月　　日以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向專利局提出。若逾上述與本所聯絡的期限，本所之服務費將會因匯率變動關係而有所調整。事關　" & custtype & "權益，尚請特加留意是幸。"
                        'Modified by Morgan 2024/5/16 EPC案調整內容,其他國家不變--郭
                        'g_WordAp.Selection.TypeText "　　六、" & custtype & "若現今即願意修正本案，本所代　" & custtype & "提出修正之費用為新台幣　　萬　　仟元整，實體審查費用為新台幣　　萬　　仟元整。倘　" & custtype & "欲續行本案，則本案可於" & m_CP07 & "前向" & IIf(nationname = "ＥＰＣ", "ＥＰＯ", nationname & "專利局") & "作出回覆，故敬請於" & m_AppDate & "以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向專利局提出。若逾上述與本所聯絡的期限，本所之服務費將會因匯率變動關係而有所調整。事關　" & custtype & "權益，尚請特加留意是幸。"
                        If pa(9) = "221" Then
                           'Modified by Morgan 2025/6/24
                           'g_WordAp.Selection.TypeText "　　六、本案應於" & m_CP07 & "前向EPO作出回覆，本所代　" & custtype & "回覆擴大檢索報告之費用為新台幣　　萬　　仟元整，實體審查費用為新台幣　　萬　　仟元整。倘　" & custtype & "欲續行本案，敬請於" & m_AppDate & "以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向專利局提出。若逾上述與本所聯絡的期限，本所之服務費將會因匯率變動關係而有所調整。事關　" & custtype & "權益，尚請特加留意是幸。"
                           g_WordAp.Selection.TypeText "　　六、本案應於" & m_CP07 & "前向EPO作出回覆，本所代　" & custtype & "回覆擴大檢索報告之費用為新台幣　　萬　　仟元整，實體審查費用為新台幣　　萬　　仟元整，會員國指定費為新台幣　　萬　　仟元整，延伸國指定費每一國為新台幣　　萬　　仟元整。倘　" & custtype & "欲續行本案，敬請於" & m_AppDate & "以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向專利局提出。若逾上述與本所聯絡的期限，本所之服務費將會因匯率變動關係而有所調整。事關　" & custtype & "權益，尚請特加留意是幸。"
                        Else
                           g_WordAp.Selection.TypeText "　　六、" & custtype & "若現今即願意修正本案，本所代　" & custtype & "提出修正之費用為新台幣　　萬　　仟元整，實體審查費用為新台幣　　萬　　仟元整。倘　" & custtype & "欲續行本案，則本案可於" & m_CP07 & "前向" & nationname & "專利局作出回覆，故敬請於" & m_AppDate & "以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向專利局提出。若逾上述與本所聯絡的期限，本所之服務費將會因匯率變動關係而有所調整。事關　" & custtype & "權益，尚請特加留意是幸。"
                        End If
                        'end 2024/5/16
                     Else
                        'Modified by Morgan 2015/9/14 "本案依法可於"->"本案可於" --郭
                        'Modified by Morgan 2021/9/3
                        'g_WordAp.Selection.TypeText "　　五、" & custtype & "若現今即願意修正本案，本所代　" & custtype & "提出修正之費用為新台幣　　萬　　仟元整，實體審查費用為新台幣　　萬　　仟元整。倘　" & custtype & "欲續行本案，則本案可於" & m_CP07 & "前向" & IIf(nationname = "ＥＰＣ", "ＥＰＯ", nationname & "專利局") & "作出回覆，故敬請於　　年　　月　　日以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向專利局提出。若逾上述與本所聯絡的期限，本所之服務費將會因匯率變動關係而有所調整。事關　" & custtype & "權益，尚請特加留意是幸。"
                        'Modified by Morgan 2024/5/16 調整內容--郭
                        'g_WordAp.Selection.TypeText "　　五、" & custtype & "若現今即願意修正本案，本所代　" & custtype & "提出修正之費用為新台幣　　萬　　仟元整，實體審查費用為新台幣　　萬　　仟元整。倘　" & custtype & "欲續行本案，則本案可於" & m_CP07 & "前向" & IIf(nationname = "ＥＰＣ", "ＥＰＯ", nationname & "專利局") & "作出回覆，故敬請於" & m_AppDate & "以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向專利局提出。若逾上述與本所聯絡的期限，本所之服務費將會因匯率變動關係而有所調整。事關　" & custtype & "權益，尚請特加留意是幸。"
                        'Modified by Morgan 2024/5/16 EPC案調整內容,其他國家不變--郭
                        'g_WordAp.Selection.TypeText "　　五、" & custtype & "若現今即願意修正本案，本所代　" & custtype & "提出修正之費用為新台幣　　萬　　仟元整，實體審查費用為新台幣　　萬　　仟元整。倘　" & custtype & "欲續行本案，則本案可於" & m_CP07 & "前向" & IIf(nationname = "ＥＰＣ", "ＥＰＯ", nationname & "專利局") & "作出回覆，故敬請於" & m_AppDate & "以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向專利局提出。若逾上述與本所聯絡的期限，本所之服務費將會因匯率變動關係而有所調整。事關　" & custtype & "權益，尚請特加留意是幸。"
                        If pa(9) = "221" Then
                           'Modified by Morgan 2025/4/21
                           'g_WordAp.Selection.TypeText "　　五、本案應於" & m_CP07 & "前向EPO作出回覆，本所代　" & custtype & "回覆擴大檢索報告之費用為新台幣　　萬　　仟元整，實體審查費用為新台幣　　萬　　仟元整。倘　" & custtype & "欲續行本案，敬請於" & m_AppDate & "以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向專利局提出。若逾上述與本所聯絡的期限，本所之服務費將會因匯率變動關係而有所調整。事關　" & custtype & "權益，尚請特加留意是幸。"
                           g_WordAp.Selection.TypeText "　　五、本案應於" & m_CP07 & "前向EPO作出回覆，本所代　" & custtype & "回覆擴大檢索報告之費用為新台幣　　萬　　仟元整，實體審查費用為新台幣　　萬　　仟元整，會員國指定費為新台幣　　萬　　仟元整，延伸國指定費每一國為新台幣　　萬　　仟元整。倘　" & custtype & "欲續行本案，敬請於" & m_AppDate & "以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向專利局提出。若逾上述與本所聯絡的期限，本所之服務費將會因匯率變動關係而有所調整。事關　" & custtype & "權益，尚請特加留意是幸。"
                        Else
                           g_WordAp.Selection.TypeText "　　五、" & custtype & "若現今即願意修正本案，本所代　" & custtype & "提出修正之費用為新台幣　　萬　　仟元整，實體審查費用為新台幣　　萬　　仟元整。倘　" & custtype & "欲續行本案，則本案可於" & m_CP07 & "前向" & nationname & "專利局作出回覆，故敬請於" & m_AppDate & "以前與本所聯絡，以便共商其中細節而能及時寄交代理人以確保本案能在法定期限之前向專利局提出。若逾上述與本所聯絡的期限，本所之服務費將會因匯率變動關係而有所調整。事關　" & custtype & "權益，尚請特加留意是幸。"
                        End If
                        'end 2024/5/16
                     End If
                     If rsA.State <> adStateClosed Then rsA.Close
                     Set rsA = Nothing
                     '2015/4/14 END
                     
                  '2008/12/1 ADD BY SONIA P案告知代理人,函知客戶檢索報告用
                  '2014/6/4加註"17"專利權評價報告(內容同14檢索報告,若檢索報告有修改,"17"也要改)  *********
                  ElseIf pa(1) = "P" Then
                     StrSQLa = "Select CF10 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='421' "
                     rsA.CursorLocation = adUseClient
                     rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                     g_WordAp.Selection.TypeText "　　一、關於　" & custtype & "前委託本所代理向" & nationname & "申請之" & AppNo & "「" & CASENAME & "」" & casetype4 & casetype1 & casetype2 & "檢索報告(本所案號" & CaseNo & ")，頃接獲代理人轉來" & rsA.Fields(0).Value & casetype1 & casetype2 & "檢索報告。"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　二、茲隨函附上本件檢索報告(含引證資料)，敬請存查。"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　三、本所意見："
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　根據" & rsA.Fields(0).Value & "的檢索報告內容，審查員共引用了相關的引證資料計有　　等　件；"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　上述　件引證資料，僅揭露出本案權利要求的部份技術內容或現有技術一部份的內容，亦即各引證資料並沒有完全的將專利技術完全公開，因此，認為本案權利要求與引證資料相較，符合具有新穎性及創造性的規定。"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　故檢索報告的初步結論，認為權利要求　至　符合專利法第二十二條有關新穎性及創造性的規定。"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　故檢索報告的初步結論，認為權利要求　至　均不符合專利法第二十二條有關創造性的規定。若　" & custtype & "往後欲主張實用新型專利權，應特別留意。"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　四、以上分析意見僅供參考，如對前述內容有任何疑問，請隨時洽詢本所，本所將竭盡全力為　" & custtype & "服務。"
                     If rsA.State <> adStateClosed Then rsA.Close
                     Set rsA = Nothing
                  End If
                  '2008/12/1 END

               '2009/3/3 ADD BY SONIA
               Case "15"    'PCT國際初步審查報告1216
                  '抓最早優先權日或申請日+30個月
                  strFirstPriDate = PUB_GetFirstPriDate(pa)
                  '法定期限=申請日(最早優先權日)+30個月
                  If strFirstPriDate <> "" Then
                     m_CP07 = CompDate(1, 30, TransDate(strFirstPriDate, 2))
                  Else
                     m_CP07 = CompDate(1, 30, TransDate(pa(10), 2))
                  End If
                  m_CP07 = Mid(m_CP07, 1, 4) & "年" & Mid(m_CP07, 5, 2) & "月" & Mid(m_CP07, 7, 2) & "日"
                  g_WordAp.Selection.TypeText "　　一、　" & custtype & "前委託本所辦理之" & AppNo & "「" & CASENAME & "」" & nationname & casetype4 & casetype1 & casetype2 & casetype3 & "案(本所案號" & CaseNo & ")，頃接代理人轉來本案之專利性國際初步報告。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　二、茲隨函附上專利性國際初步報告與前案影本乙份，敬請存查。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　三、審查意見："
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　四、本所意見："
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　(註：有部份項次不具專利性)"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　截至目前為止，本案的國際申請階段已全部完成，依法規定應於申請日(有優先權日者，以優先權日為準)起算30個月內進入各國國家階段，因此，本案最遲應於" & m_CP07 & "前完成進入各國國家階段的手續。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　由於專利性國際初步報告中指出本案之申請專利範圍第  至  項不具新穎性、第  至  項不具創造性，故　" & custtype & "在日後進入各國國家階段的同時，可以另外提出主動修正申請，以根據目前國際審查單位的審查意見來修正本案，如此讓各國審查委員依新修正的申請專利範圍進行審查，使本案有機會先一步避開專利性國際初步報告中所引用的前案。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　另在進入各國國家階段時可選擇暫不修正，於進入各國國家階段後再依據各國的審定書作出回覆，屆時各國審查委員將參考專利性國際初步報告的內容，依據其判斷進行審查。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "=========================================================="
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　四、本所意見："
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　(註：全部項次均具有專利要件)"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　截至目前為止，本案的國際申請階段已全部完成，依法規定應於申請日(有優先權日者，以優先權日為準)起算30個月內進入各國國家階段，因此，本案最遲應於" & m_CP07 & "前完成進入各國國家階段的手續。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　五、以上意見僅供參考，如對前述內容有任何疑問，請隨時洽詢本所，本所將竭盡全力為　" & custtype & "服務。"
               '2009/3/3 END
               
               'Add By Sindy 2013/3/29
               Case "16"      '台灣舉發及答辯審定書
                                                      
                  '2013/4/26 modify by sonia 案件名稱欄:卷宗性質='1'者抓pa05,非'1'者才抓cp37; 不抓casetype3固定為舉發案
                  'StrSQLa = "Select C1.CP09,C2.CP115 as CP115,C2.CP08 as CP08,C2.CP06 as CP06,C3.CP36 as CP36,C3.CP37 as CP37 FROM CaseProgress C1,CaseProgress C2,CaseProgress C3 WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                            " AND C1.CP10='941' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP43=C2.CP09(+) AND C2.CP10 in ('1001','1002','1503') AND C2.CP43=C3.CP09(+) AND C3.CP10 in ('803','804')"
                  'Modified by Morgan 2021/10/21 +Flg1(通知聽證1812收文號)
                  'Modified by Morgan 2022/8/24 +部分准駁1009
                  StrSQLa = "Select C1.CP09,C2.CP115 as CP115,C2.CP08 as CP08,C2.CP06 as CP06,C3.CP36 as CP36,decode(pa23,'1',pa05,C3.CP37) as CP37,C3.CP10,C4.CP09 as Flg1" & _
                            " FROM CaseProgress C1,CaseProgress C2,CaseProgress C3,PATENT,CaseProgress C4" & _
                            " WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                            " AND C1.CP10='941' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP43=C2.CP09(+) AND C2.CP10 in ('1001','1002','1009','1503') AND C2.CP43=C3.CP09(+) AND C3.CP10 in ('803','804')" & _
                            " AND C1.CP01=PA01(+) AND C1.CP02=PA02(+) AND C1.CP03=PA03(+) AND C1.CP04=PA04(+) " & _
                            " and C4.CP43(+)=C3.CP09 and C4.CP10(+)='1812'"
                  '2013/4/26 end
                  
                  'Added by Morgan 2016/3/16
                  'P臺灣案的答辯或舉發答辯的審定來函改工程師承辦(原內部收文分析)
                  'Modified by Morgan 2021/10/21 +Flg1(通知聽證1812收文號)
                  'Modified by Morgan 2022/8/24 +部分准駁1009
                  StrSQLa = StrSQLa & " union Select a.CP09,a.CP115,a.CP08,a.CP06,b.CP36,decode(pa23,'1',pa05,b.CP37) as CP37,b.cp10,C4.CP09 as Flg1" & _
                     " FROM CaseProgress a,CaseProgress b,patent,CaseProgress C4" & _
                     " WHERE a.CP01='" & pa(1) & "' AND a.CP02='" & pa(2) & "' AND a.CP03='" & pa(3) & "' AND a.CP04='" & pa(4) & "' and a.cp10 in ('1001','1002','1009','1503') and a.cp27||a.cp57 is null" & _
                     " and b.CP09(+)=a.CP43 AND b.CP10 in ('803','804') and pa01(+)=a.cp01 and pa02(+)=a.cp02 and pa03(+)=a.cp03 and pa04(+)=a.cp04" & _
                     " and C4.CP43(+)=b.CP09 and C4.CP10(+)='1812'"
                  'end 2016/3/16
                  
                  rsA.CursorLocation = adUseClient
                  rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                  If rsA.RecordCount > 0 Then
                     If IsNull(rsA.Fields("CP115").Value) Then
                        m_CP115 = "　年　月　日"
                     Else
                        'modify by sonia 2019/11/20 判斷年度格式
                        m_CP115 = IIf(bolDateType, Mid(rsA.Fields("CP115").Value, 1, 4), Mid(rsA.Fields("CP115").Value, 1, 4) - 1911) & "年" & Mid(rsA.Fields("CP115").Value, 5, 2) & "月" & Mid(rsA.Fields("CP115").Value, 7, 2) & "日"
                     End If
                     If IsNull(rsA.Fields("CP06").Value) Then
                        m_CP06 = "　年　月　日"
                        m_DueDate = "　年　月　日"
                     Else
                        'modify by sonia 2019/11/20 判斷年度格式
                        m_CP06 = IIf(bolDateType, Mid(rsA.Fields("CP06").Value, 1, 4), Mid(rsA.Fields("CP06").Value, 1, 4) - 1911) & "年" & Mid(rsA.Fields("CP06").Value, 5, 2) & "月" & Mid(rsA.Fields("CP06").Value, 7, 2) & "日"
                        '本所期限-7天
                        m_DueDT = CompDate(2, -7, "" & rsA.Fields("CP06").Value)
                        m_DueDT = PUB_GetWorkDay1(m_DueDT, True)        'add by sonia 2020/12/24 取工作天
                        'modify by sonia 2019/11/20 判斷年度格式
                        m_DueDate = IIf(bolDateType, Mid(m_DueDT, 1, 4), Mid(m_DueDT, 1, 4) - 1911) & "年" & Mid(m_DueDT, 5, 2) & "月" & Mid(m_DueDT, 7, 2) & "日"
                     End If
                     strRefCP10 = "" & rsA.Fields("CP10").Value 'Added by Morgan 2015/5/26
                  Else
                     m_CP115 = "　年　月　日"
                     m_CP06 = "　年　月　日"
                     m_DueDate = "　年　月　日"
                  End If
                  '2013/4/26 modify by sonia 案件名稱欄:卷宗性質='1'者抓pa05,非'1'者才抓cp37,已在上面抓資料語法中控制; 不抓casetype3固定為舉發案
                  'g_WordAp.Selection.TypeText "　　" & custtype & "委由本所辦理之" & rsA.Fields("CP36").Value & "「" & rsA.Fields("CP37").Value & "」" & casetype4 & casetype1 & casetype2 & casetype3 & "案(本所案號" & CaseNo & ")，頃接經濟部智慧財產局(以下稱智慧局)於" & m_CP115 & "發出的" & rsA.Fields("CP08") & "專利舉發審定書(如附件)。"
                  'Modified by Morgan 2015/5/26
                  'g_WordAp.Selection.TypeText "　　" & custtype & "委由本所辦理之" & rsA.Fields("CP36").Value & "「" & rsA.Fields("CP37").Value & "」" & casetype4 & casetype1 & casetype2 & "舉發案(本所案號" & CaseNo & ")，頃接經濟部智慧財產局(以下稱智慧局)於" & m_CP115 & "發出的" & rsA.Fields("CP08") & "專利舉發審定書(如附件)。"
                  'Modified by Morgan 2020/4/22 +補第..號 --郭 Ex:P-124144
                  'Modified by Lydia 2020/10/15 委由=>委託
                  g_WordAp.Selection.TypeText "　　" & custtype & "委託本所辦理之第" & rsA.Fields("CP36").Value & "號「" & rsA.Fields("CP37").Value & "」" & casetype4 & casetype1 & casetype2 & IIf(strRefCP10 = "804", "舉發答辯", "舉發") & "案(本所案號" & CaseNo & ")，頃接經濟部智慧財產局(以下稱智慧局)於" & m_CP115 & "發出的" & rsA.Fields("CP08") & "專利舉發審定書(如附件)。"
                  'end 2015/5/26
                  '2013/4/26 end
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　審定主文："
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　〔註：本段請記載審定主文的全部內容，完成後刪除本註解。〕"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　〔申請專利範圍的審查版本：若被舉發案曾更正申請專利範圍，請選擇其中一段內容說明，若未曾更正，則本段省略，完成後刪除本註解。〕"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　智慧局在審定理由中，審定不准予（我方或對造）於　　年　月　日所提出之申請專利範圍更正並依原公告內容進行審查。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　智慧局在審定理由中，審定准予（我方或對造）於　　年　月　日所提出之申請專利範圍更正並依更正後的內容進行審查。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　〔註：以下兩種內容，請配合審定主文為部份或全部請求項次為舉發成立或不成立予以選擇或同時運用，完成後刪除本註解。〕"
                  g_WordAp.Selection.TypeParagraph
                  
                  'Added by Morgan 2021/10/21 有通知聽證1812
                  If Not IsNull(rsA.Fields("Flg1")) Then
                     g_WordAp.Selection.TypeText "　　關於申請專利範圍請求項　　至　　，智慧局採信我方提出的" & IIf(strRefCP10 = "804", "理由", "爭點及理由") & "，為「舉發成立應予撤銷或舉發不成立」之審定。惟對造如不服此項審定，可於審定書送達次日起二個月內向智慧財產及商業法院提出行政訴訟，倘對造逾期未提出行政訴訟，則此部分之審定即告確定。"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　關於申請專利範圍請求項　　至　　，智慧局為「舉發不成立或舉發成立應予撤銷」之審定。如　" & custtype & "不服此項之審定，因本案係經聽證作成之行政處分，依行政程序法第109條規定本案可免除訴願及其先行程序，可依法於" & m_CP06 & "以前提出行政訴訟，若欲辦理，請　" & custtype & "提早於" & m_DueDate & "前通知本所，以便在法定期限之前向智慧財產及商業法院提出。"
                  Else
                  'end 2021/10/21
                  
                     'Modified by Morgan 2015/5/26
                     'g_WordAp.Selection.TypeText "　　關於申請專利範圍請求項　　至　　，智慧局採信我方提出的爭點及理由，為「舉發成立應予撤銷或舉發不成立」之審定。惟對造如不服此項審定，可於審定書送達次日起30日內向經濟部提出訴願，倘對造逾期未提出訴願，則此部分之審定即告確定。"
                     g_WordAp.Selection.TypeText "　　關於申請專利範圍請求項　　至　　，智慧局採信我方提出的" & IIf(strRefCP10 = "804", "理由", "爭點及理由") & "，為「舉發成立應予撤銷或舉發不成立」之審定。惟對造如不服此項審定，可於審定書送達次日起30日內向經濟部提出訴願，倘對造逾期未提出訴願，則此部分之審定即告確定。"
                     'end 2015/5/526
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　關於申請專利範圍請求項　　至　　，智慧局為「舉發不成立或舉發成立應予撤銷」之審定。如　" & custtype & "不服此項之審定，可依法於" & m_CP06 & "以前提出訴願，若欲辦理，請　" & custtype & "提早於" & m_DueDate & "前通知本所，以便在法定期限之前向經濟部提出。"
                  
                  End If 'Added by Morgan 2021/10/21
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　若有任何問題，請不吝隨時賜教，本所將竭誠為　" & custtype & "提供最佳服務。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　隨函附送本案之專利舉發審定書一份，敬請查收備查。"
                  If rsA.State <> adStateClosed Then rsA.Close
                  Set rsA = Nothing
               '2013/3/29 End
               'ADD BY SONIA 2014/6/4
               Case "17"      '專利權評價報告(內容同14檢索報告,若檢索報告有修改,此處也要改)P-105534
                  If pa(1) = "P" Then
                     StrSQLa = "Select CF10 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='423' "
                     rsA.CursorLocation = adUseClient
                     rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                     g_WordAp.Selection.TypeText "　　一、關於　" & custtype & "前委託本所代理向" & nationname & "申請之" & AppNo & "「" & CASENAME & "」" & casetype4 & casetype1 & "專利權評價報告(本所案號" & CaseNo & ")，頃接獲代理人轉來" & rsA.Fields(0).Value & casetype1 & "專利權評價報告。"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　二、茲隨函附上本件評價報告(含引證資料)，敬請存查。"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　三、本所意見："
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　根據" & rsA.Fields(0).Value & "的評價報告內容，審查員共引用了相關的引證資料計有　　等　件；"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　上述　件引證資料，僅揭露出本案權利要求的部份技術內容或現有技術一部份的內容，亦即各引證資料並沒有完全的將專利技術完全公開，因此，認為本案權利要求與引證資料相較，符合具有新穎性及創造性的規定。"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　故評價報告的初步結論，認為權利要求　至　符合專利法第二十二條有關新穎性及創造性的規定。"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　故評價報告的初步結論，認為權利要求　至　均不符合專利法第二十二條有關創造性的規定。若　" & custtype & "往後欲主張實用新型專利權，應特別留意。"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　四、以上分析意見僅供參考，如對前述內容有任何疑問，請隨時洽詢本所，本所將竭盡全力為　" & custtype & "服務。"
                     If rsA.State <> adStateClosed Then rsA.Close
                     Set rsA = Nothing
                  End If
               'END 2014/6/4
            End Select
         '2008/9/18 add by sonia
         Case "2"             '核駁前先行通知1202
            'Modify By Sindy 2012/4/23 +CP06
            StrSQLa = "Select CP07,CP08,CP36,NVL(NVL(CP37,CP38),CP39),NVL(NVL(TG06||TG15,TG07||TG16),TG08||TG17),CP06 FROM CaseProgress,TMGOODS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                      IIf(m_T727CP43No <> "", " AND CP09='" & m_T727CP43No & "' ", " AND CP10='1202' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' ") & " AND CP01=TG01(+) AND CP02=TG02(+) AND CP03=TG03(+) AND CP04=TG04(+) "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               'modify by sonia 2019/11/20 判斷年度格式
               m_CP07 = IIf(bolDateType, Mid(rsA.Fields(0).Value, 1, 4), Mid(rsA.Fields(0).Value, 1, 4) - 1911) & "年" & Mid(rsA.Fields(0).Value, 5, 2) & "月" & Mid(rsA.Fields(0).Value, 7, 2) & "日"
               'Add By Sindy 2012/4/23
               'modify by sonia 2019/11/20 判斷年度格式
               m_CP06 = IIf(bolDateType, Mid(rsA.Fields("CP06").Value, 1, 4), Mid(rsA.Fields("CP06").Value, 1, 4) - 1911) & "年" & Mid(rsA.Fields("CP06").Value, 5, 2) & "月" & Mid(rsA.Fields("CP06").Value, 7, 2) & "日"
               '本所期限-3工作天
               m_DueDT = CompWorkDay(3, rsA.Fields("CP06").Value, 1)
               'modify by sonia 2019/11/20 判斷年度格式
               m_DueDate = IIf(bolDateType, Mid(m_DueDT, 1, 4), Mid(m_DueDT, 1, 4) - 1911) & "年" & Mid(m_DueDT, 5, 2) & "月" & Mid(m_DueDT, 7, 2) & "日"
               '2012/4/23 End
               'add by sonia 2021/4/20 多商品類別的商品也要都抓T-231208,下面定稿內容程式rsA.Fields(4)換成m_TMGoods
               m_TMGoods = ""
               Do While Not rsA.EOF
                  If m_TMGoods <> "" Then m_TMGoods = m_TMGoods & "；"
                  m_TMGoods = m_TMGoods & rsA.Fields(4)
                  rsA.MoveNext
               Loop
               rsA.MoveFirst
               'end 2021/4/20
            
               'Modify By Sindy 2024/3/26 依對造號數判斷是註冊還是申請字樣
               'IIf(Len(Trim("" & rsA.Fields(2))) = 8, "註冊", "申請")
               strCP36Kind = "申請"
               If Trim("" & rsA.Fields(2)) <> "" Then
                  tmpArr = Split(Trim("" & rsA.Fields(2)), ",")
                  strExc(9) = tmpArr(LBound(tmpArr))
                  If Len(strExc(9)) = 8 Then
                     strCP36Kind = "註冊"
                  End If
               End If
               '2024/3/26 END
            End If
            
            If pa(10) = "000" Then    '台->台
               'Modified by Lydia 2020/10/15 委由=>委託
               g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之申請" & AppNo & "「" & CASENAME & "」" & casetype4 & "（第" & pa(9) & "類）商標" & casetype3 & "乙案，經智慧財產局初步審查後，認有違反商標法" & caselaw & "規定之嫌，茲檢附智慧局" & rsA.Fields(1) & "核駁理由先行通知書乙紙，請查照。"
               g_WordAp.Selection.TypeParagraph
               'Modify By Sindy 2012/4/23
               'g_WordAp.Selection.TypeText "二、本案依法應於" & m_CP07 & "以前提出意見書，逾期本件商標將遭核駁處分。有關本案本所以為："
               g_WordAp.Selection.TypeText "二、本案依法應於" & m_CP06 & "以前提出意見書，逾期本件商標將遭核駁處分，　" & custtype & "若有意續行，請於" & m_DueDate & "以前與本所聯繫，以共商相關事宜。有關本案本所以為："
               g_WordAp.Selection.TypeParagraph
               'Modify By Sindy 2012/7/11 101年7月1日商標修法
               Select Case m_Combo8
                  '2008/11/19 add by sonia
                  Case "2C"      'T核駁前先行通知1202第23條第1項第1款(AXX)及第23條第1項第2款(BXX)及第23條第1項第11款(FXX)
                     g_WordAp.Selection.TypeText "　　智慧局認為商標圖樣上之「" & CASENAME & "」，有「」之意，以之作為商標，不足以使商品購買人認識其為表彰商品來源之標識，並得藉以與他人之商品相區別，欠缺識別性；又以「" & CASENAME & "」指定使用於「" & m_TMGoods & "」商品，有商品說明之虞；指定使用於「」等商品，有使公眾對其商品之內容、性質產生誤信誤認之虞。依現行之審查基準，　" & custtype & "可聲明圖樣中之「" & CASENAME & "」部分不在專用之列，並將指定商品減縮為「」；或刪除圖樣中之「" & CASENAME & "」，保留原指定之所有商品，以爭取本案之註冊。"
                  '2008/11/19 end
                  '2010/10/7 add by sonia
                  Case "2D"      'T核駁前先行通知1202第23條第1項第1款(AXX)及第23條第1項第2款(BXX)及第23條第1項第13款(HXX)
                     g_WordAp.Selection.TypeText "　　智慧局認為以「" & CASENAME & "」作為商標之一部分，指定使用於「" & m_TMGoods & "」商品，不足以使相關消費者認識其為表彰商品來源之標識，並得藉以與他人之商品相區別，欠缺識別性；此外本件商標圖樣中之「" & CASENAME & "」與" & strCP36Kind & "第" & rsA.Fields(2) & "號「" & rsA.Fields(3) & "」商標圖樣相較，兩者構圖意匠相彷彿，復指定使用於相同或類似商品，有使一般消費者產生混淆誤認之虞，應屬近似商標。　" & custtype & "可取得經前揭商標所有人同意申請之證明文件，或刪除與前揭商標構成近似之「" & m_TMGoods & "」商品，僅保留「」商品，並聲明圖樣中之「」不在專用之列，爭取之；否則以重新設計圖樣，另案申請為宜。"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　　" & custtype & "若已使用該商標，可據實陳述意見，舉證無使消費者混淆誤認之虞之具體事證，並聲明「" & CASENAME & "」不在專用之列，以爭取之；否則以更改商標圖樣，另案申請為宜。"
                  '2010/10/7 end
                  'Modify By Sindy 2012/9/6 修改定稿內容
                  Case "21"      'T核駁前先行通知1202第29條第1項第1款(AXX)及第29條第1項第3款(BXX)
                     g_WordAp.Selection.TypeText "　　智慧局認為商標圖樣上之「" & CASENAME & "」，予消費者的認知，僅為商品或服務之說明或為……之描述用語，以之作為商標，指定使用於「" & m_TMGoods & "」等商品，缺乏商標應有的顯著特徵與識別商品或服務來源之功能，並得藉以與他人之商品或服務相區別，欠缺識別性。惟　" & custtype & "若有意爭取，可說明本件商標之設計原由及實際使用態樣，並提出廣告等資料，證明經長期且廣泛使用，該商標在交易市場上已成為指定商品或服務之識別標識，已具識別性，以爭取之；否則以重新設計圖樣，另案申請為宜。"
                  '2012/9/6 End
                  'Add By Sindy 2012/9/6
                  Case "2M"      'T核駁前先行通知1202第29條第1項第2款(AXX)及第29條第1項第3款(CXX)
                     g_WordAp.Selection.TypeText "　　智慧局認為以「" & CASENAME & "」作為商標，指定使用於「" & m_TMGoods & "」等商品，為所指定商品之通用名稱或為……之描述用語，缺乏商標應有的顯著特徵與識別商品或服務來源之功能，並得藉以與他人之商品或服務相區別，欠缺識別性。惟　" & custtype & "若有意爭取，可說明本件商標之設計原由及實際使用態樣，並提出廣告等資料，證明經長期且廣泛使用，該商標在交易市場上已成為指定商品或服務之識別標識，已具識別性，以爭取之；否則以重新設計圖樣，另案申請為宜。"
                  '2012/9/6 End
'                  '2011/2/25 Add By Sindy
'                  Case "2E"      'T核駁前先行通知1202第23條第1項第1款(AXX)及第13款(HXX)
'                     g_WordAp.Selection.TypeText "　　智慧局認為以「" & CASENAME & "」作為商標之一部分，指定使用於「" & rsA.Fields(4) & "」商品，不足以使相關消費者認識其為表彰商品來源之標識，並得藉以與他人之商品相區別，欠缺識別性；此外本件商標圖樣中之「" & CASENAME & "」與註冊第" & rsA.Fields(2) & "號「" & rsA.Fields(3) & "」商標圖樣相較，兩者構圖意匠相彷彿，復指定使用於相同或類似商品，有使一般消費者產生混淆誤認之虞，應屬近似商標。　" & custtype & "可取得經前揭商標所有人同意申請之證明文件，或刪除與前揭商標構成近似之「" & rsA.Fields(4) & "」商品，僅保留「」商品，並聲明圖樣中之「」不在專用之列，爭取之；否則以重新設計圖樣，另案申請為宜。"
'                     g_WordAp.Selection.TypeParagraph
'                     g_WordAp.Selection.TypeText "　　　" & custtype & "若已使用該商標，可據實陳述意見，舉證無使消費者混淆誤認之虞之具體事證，並聲明「" & CASENAME & "」不在專用之列，以爭取之；否則以更改商標圖樣，另案申請為宜。"
'                  '2011/2/25 End
'
'                  Case "22"      'T核駁前先行通知1202第23條第1項第2款(BXX)及第23條第1項第11款(FXX)
'                     g_WordAp.Selection.TypeText "　　智慧局認為以「" & CASENAME & "」作為商標，指定使用於「" & rsA.Fields(4) & "」商品，為商品功用之說明；又，指定使用於「」商品，有使公眾對其商品之性質、品質發生誤認誤信之虞。　" & custtype & "若有意爭取，可聲明圖樣中「" & CASENAME & "」不在專用之列，並刪除「" & rsA.Fields(4) & "」商品，僅保留「」等商品，本件商標方有獲准註冊之機會。"
'                     g_WordAp.Selection.TypeParagraph
'                     g_WordAp.Selection.TypeText "　　" & custtype & "若有意爭取，可刪除「」商品，並據實陳述意見，檢附實際使用之態樣，證明業經長期且廣泛使用，該商標在交易市場上已成為指定商品之識別標識，足資做為消費者辨識商品來源之依據，已具識別性，以爭取之。"
'                  Case "23"      'T核駁前先行通知1202第23條第1項第2款(BXX)及第23條第1項第13款(HXX)
'                     g_WordAp.Selection.TypeText "　　智慧局認為以「" & CASENAME & "」作為商標之一部分，指定使用於「" & rsA.Fields(4) & "」商品，為所指定商品之說明；此外本件商標圖樣中之「」與註冊第" & rsA.Fields(2) & "號「" & rsA.Fields(3) & "」商標圖樣相較，　，兩者構圖意匠相彷彿，復指定使用於相同或類似商品，有使一般商品購買人產生混淆誤認之虞，應屬近似商標。　" & custtype & "可取得經前揭商標所有人同意申請之證明文件，或刪除與前揭商標構成近似之「" & rsA.Fields(4) & "」商品，僅保留「」商品，並聲明圖樣中之「" & CASENAME & "」不在專用之列，爭取之；否則以重新設計圖樣，另案申請為宜。"
'                     g_WordAp.Selection.TypeParagraph
'                     g_WordAp.Selection.TypeText "　　" & custtype & "若已使用該商標，可據實陳述意見，舉證無使消費者混淆誤認之虞之具體事證，並聲明「" & CASENAME & "」不在專用之列，以爭取之；否則以更改商標圖樣，另案申請為宜。"
'                  Case "24"      'T核駁前先行通知1202第23條第1項第11款(FXX)及第23條第1項第13款(HXX)
'                     g_WordAp.Selection.TypeText "　　智慧局認為以「" & CASENAME & "」作為商標之一部分，指定使用於「" & rsA.Fields(4) & "等商品，有使公眾誤認誤信商品性質、品質之虞；此外，本件商標之「」與註冊第" & rsA.Fields(2) & "號「" & rsA.Fields(3) & "」商標相較，　，復指定使用於相同或類似商品，應屬近似商標。　" & custtype & "若已使用該商標，可據實陳述意見，舉證無使消費者混淆誤認之虞，並刪除圖樣上「" & CASENAME & "」部分，以爭取之。"
'                  Case "25"      'T核駁前先行通知1202第23條第1項第1款(AXX)
'                     g_WordAp.Selection.TypeText "　　智慧局認為商標圖樣上之「" & CASENAME & "」，予人印象為一般描述用語，以之作為商標，指定使用於「" & rsA.Fields(4) & "」等商品，不足以使商品購買人認識其為表彰商品來源之標識，並得藉以與他人之商品相區別，欠缺識別性。惟　" & custtype & "若有意爭取，可說明本件商標之設計原由及實際使用態樣，並提出廣告等資料，證明經長期且廣泛使用，該商標在交易市場上已成為指定商品之識別標識，足資做為消費者辨識商品來源之依據，已具識別性，以爭取之；或同意聲明圖樣中「" & CASENAME & "」部分不在專用之列，本件商標方有獲准註冊之機會。"
'                  Case "26"      'T核駁前先行通知1202'第23條第1項第2款(BXX)
'                     g_WordAp.Selection.TypeText "　　智慧局認為以「" & CASENAME & "」作為商標之一部分，指定使用於「" & rsA.Fields(4) & "」等商品，有說明商品之虞，無法使一般消費者認識其為表彰商品來源之標識，並得藉以與他人商品相區別，欠缺顯著性。惟　" & custtype & "若有意爭取，可說明本件商標之設計原由及實際使用態樣，並提出廣告等資料，證明經長期且廣泛使用，該商標在交易市場上已成為指定商品之識別標識，足資做為消費者辨識商品來源之依據，已具識別性，以爭取之；或同意聲明圖樣中「" & CASENAME & "」部分不在專用之列，本件商標方可獲准註冊。"
'                  Case "27"      'T核駁前先行通知1202第23條第1項第11款(FXX)
'                     g_WordAp.Selection.TypeText "　　智慧局認為以「" & CASENAME & "」作為商標之一部分，指定使用於「" & rsA.Fields(4) & "」商品，有使相關消費者對其表彰商品之性質產生誤認誤信之虞。本案建議　" & custtype & "可減縮「」等商品；或同意刪除圖樣中之「" & CASENAME & "」部分，以爭取本案之獲准註冊。"
'                     g_WordAp.Selection.TypeParagraph
'                     g_WordAp.Selection.TypeText "　　智慧局認為「" & CASENAME & "」係；以之作為商標，指定使用於「" & rsA.Fields(4) & "」商品，有使公眾對其表彰服務之提供者發生誤認誤信之虞。　" & custtype & "可據實陳述意見，具體說明本件商標之設計原由，並提出無使消費者誤認商品性質之具體事證，爭取之；否則以重新設計圖樣，另案申請為宜。"
'                  Case "28"      'T核駁前先行通知1202第23條第1項第12款(GXX)
'                     g_WordAp.Selection.TypeText "　　按「商標相同或近似於他人著名商標或標章，有致相關公眾混淆誤認之虞，或有減損著名商標或標章之識別性或信譽之虞者，不得註冊」為商標法" & caselaw & "所明定。智慧局認為本件商標圖樣上之，有使公眾產生混淆誤認之虞，或有減損著名商標之識別性之虞，此業經智慧局以　　認定在案。本案　" & custtype & "可據實陳述意見，舉證無造成消費者混淆誤認或減損著名商標識別性之虞之具體事證，爭取之；否則以更改商標圖樣，另案申請為宜。"
'                  Case "29"      'T核駁前先行通知1202第23條第1項第13款(HXX)
'                     g_WordAp.Selection.TypeText "　　智慧局認為本件商標圖樣上之「" & CASENAME & "」與註冊第" & rsA.Fields(2) & "號「" & rsA.Fields(3) & "」商標近似，復指定使用於同一或類似商品，有使一般消費者產生混淆誤認之虞，應屬近似商標。　" & custtype & "若有意爭取，可據實陳述意見，檢附實際使用之態樣，說明無造成消費者混淆誤認之虞；或刪除與前揭商標構成近似之「" & rsA.Fields(4) & "」商品，僅保留「」商品，以爭取本案之註冊。"
'                     g_WordAp.Selection.TypeParagraph
'                     g_WordAp.Selection.TypeText "　　　" & custtype & "若已使用該商標，可據實陳述意見，舉證無使消費者混淆誤認之虞之具體事證，或取得經前揭商標所有人同意申請之證明文件，爭取之；否則以重新設計圖樣，另案申請為宜。"
'                  Case "2A"      'T核駁前先行通知1202第23條第1項第14款(IXX)
'                     g_WordAp.Selection.TypeText "　　按「商標相同或近似於他人先使用於同一或類似商品或服務之商標，而申請人因與該他人間具有契約、地緣、業務往來或其他關係，知悉他人商標存在者，不得註冊」，為商標法" & caselaw & "所明定。智慧局認為本件商標之「" & CASENAME & "」與　近似， ，而有違前揭法條之規定。　" & custtype & "可據實陳述意見，舉證無造成混淆誤認之虞；或檢送經前揭商標申請人同意申請之證明文件，爭取之；否則以重新設計圖樣，另案申請為宜。"
'                     g_WordAp.Selection.TypeParagraph
'                  Case "2B"      'T核駁前先行通知1202第23條第1項第15款(JXX)
'                     g_WordAp.Selection.TypeText "　　按「商標有他人之肖像或著名之姓名、藝名、筆名、字號者，不得註冊」為商標法" & caselaw & "之所明定。智慧局認為本件商標圖樣係以真實人物照片為構圖，有前揭法條之適用。　" & custtype & "可檢送經該他人同意申請註冊之證明文件，本件商標方有核准註冊之機會。"
'                     g_WordAp.Selection.TypeParagraph
'                     g_WordAp.Selection.TypeText "　　按「商標有他人之肖像或著名之姓名、藝名、筆名、字號者，不得註冊」為商標法" & caselaw & "之所明定。智慧局認為「" & CASENAME & "」係　，以「" & CASENAME & "」作為商標之一部分，指定使用於「" & rsA.Fields(4) & "」等商品，有前揭法條之適用。　" & custtype & "可據實陳述意見，具體說明本件商標之設計原由及實際使用態樣，爭取之，否則以重新設計商標圖樣，另案申請為宜。"
                  Case "2E"      'T核駁前先行通知1202第29條第1項第3款(AXX)及第30條第1項第10款(HXX)
                     g_WordAp.Selection.TypeText "　　智慧局認為以「" & CASENAME & "」作為商標，指定使用於「" & m_TMGoods & "」等商品，為指定商品或服務之通用標章或名稱，缺乏商標應有的顯著特徵與識別商品或服務來源之功能；此外本件商標圖樣中之「" & CASENAME & "」與" & strCP36Kind & "第" & rsA.Fields(2) & "號「" & rsA.Fields(3) & "」商標圖樣相較，……，兩者構圖意匠相彷彿，復指定使用於相同或類似商品或服務，有使一般消費者產生混淆誤認之虞，應屬近似商標。　" & custtype & "若有意爭取，可提出經長期且廣泛使用，該商標在交易市場上已成為指定商品或服務之識別標識，已具識別性，並舉證無使消費者混淆誤認之虞之具體事證，以爭取之；否則以重新設計圖樣，另案申請為宜。"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　　" & custtype & "若有意爭取，可同意刪除與前揭商標構成近似之「" & m_TMGoods & "」商品，僅保留「」商品，並提出經長期且廣泛使用，該商標在交易市場上已成為指定商品或服務之識別標識，已具識別性，以爭取之。"
                  Case "22"      'T核駁前先行通知1202第29條第1項第1款(BXX)及第30條第1項第8款(FXX)
                     g_WordAp.Selection.TypeText "　　智慧局認為商標圖樣上之「" & CASENAME & "」，予消費者的認知，僅為……之描述用語，以之作為商標，指定使用於「" & m_TMGoods & "」等商品，缺乏商標應有的顯著特徵與識別商品或服務來源之功能；又，指定使用於「」商品，有使公眾對其商品之性質、品質發生誤認誤信之虞。　" & custtype & "若有意爭取，可刪除「」商品，並據實陳述意見，檢附實際使用之態樣，證明業經長期且廣泛使用，該商標在交易市場上已成為指定商品之識別標識，足資做為消費者辨識商品來源之依據，已具識別性，以爭取之。"
                  Case "23"      'T核駁前先行通知1202第29條第1項第1款(BXX)及第30條第1項第10款(HXX)
                     g_WordAp.Selection.TypeText "　　智慧局認為商標圖樣上之「" & CASENAME & "」，予消費者的認知，僅為……之描述用語，以之作為商標，指定使用於「" & m_TMGoods & "」等商品，缺乏商標應有的顯著特徵與識別商品或服務來源之功能；此外本件商標圖樣中之「" & CASENAME & "」與" & strCP36Kind & "第" & rsA.Fields(2) & "號「" & rsA.Fields(3) & "」商標圖樣相較，……，兩者構圖意匠相彷彿，復指定使用於相同或類似商品或服務，有使一般消費者產生混淆誤認之虞，應屬近似商標。　" & custtype & "若有意爭取，可提出經長期且廣泛使用，該商標在交易市場上已成為指定商品或服務之識別標識，已具識別性，並舉證無使消費者混淆誤認之虞之具體事證，以爭取之；否則以重新設計圖樣，另案申請為宜。"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　　" & custtype & "若有意爭取，可同意刪除與前揭商標構成近似之「" & m_TMGoods & "」商品，僅保留「」商品，並提出經長期且廣泛使用，該商標在交易市場上已成為指定商品或服務之識別標識，已具識別性，以爭取之。"
                  Case "24"      'T核駁前先行通知1202第30條第1項第8款(FXX)及第30條第1項第10款(HXX)
                     g_WordAp.Selection.TypeText "　　智慧局認為以「" & CASENAME & "」作為商標之一部分，指定使用於「" & m_TMGoods & "」等商品，有使公眾誤認誤信商品性質、品質之虞；此外，本件商標之「" & CASENAME & "」與" & strCP36Kind & "第" & rsA.Fields(2) & "號「" & rsA.Fields(3) & "」商標相較，……，復指定使用於相同或類似之商品或服務，應屬近似商標。　" & custtype & "若已使用該商標，可據實陳述意見，舉證無使消費者混淆誤認之虞，並刪除圖樣上「" & CASENAME & "」部分，以爭取之。"
                  Case "2G"      'T核駁前先行通知1202第29條第1項第2款(CXX)及第30條第1項第8款(FXX)
                     g_WordAp.Selection.TypeText "　　智慧局認為以「" & CASENAME & "」作為商標，指定使用於「" & m_TMGoods & "」等商品，為指定商品或服務之通用標章或名稱，缺乏商標應有的顯著特徵與識別商品或服務來源之功能；又，指定使用於「」商品，有使公眾對其商品之性質、品質發生誤認誤信之虞。　" & custtype & "若有意爭取，可刪除「」商品，並據實陳述意見，檢附實際使用之態樣，證明業經長期且廣泛使用，該商標在交易市場上已成為指定商品之識別標識，足資做為消費者辨識商品來源之依據，已具識別性，以爭取之。"
                  Case "2H"      'T核駁前先行通知1202第29條第1項第3款(AXX)及第30條第1項第8款(FXX)
                     g_WordAp.Selection.TypeText "　　智慧局認為商標圖樣上之「" & CASENAME & "」，予消費者的認知，僅為……之描述用語，以之作為商標，指定使用於「" & m_TMGoods & "」等商品，缺乏商標應有的顯著特徵與識別商品或服務來源之功能；又，指定使用於「」商品，有使公眾對其商品之性質、品質發生誤認誤信之虞。　" & custtype & "若有意爭取，可刪除「」商品，並據實陳述意見，檢附實際使用之態樣，證明業經長期且廣泛使用，該商標在交易市場上已成為指定商品之識別標識，足資做為消費者辨識商品來源之依據，已具識別性，以爭取之。"
                  Case "2J"      'T核駁前先行通知1202第29條第3項(MXX)及第30條第1項第8款(FXX)
                     g_WordAp.Selection.TypeText "　　智慧局認為以「" & CASENAME & "」作為商標，指定使用於「" & m_TMGoods & "」商品，為商品品質、特性之說明；又，指定使用於「」商品，有使公眾對其商品之性質、品質發生誤認誤信之虞。　" & custtype & "若有意爭取，"
                     g_WordAp.Selection.Font.Bold = True
                     g_WordAp.Selection.TypeText "可聲明本件商標不就「" & CASENAME & "」主張商標權，並刪除「」商品"
                     g_WordAp.Selection.Font.Bold = False
                     g_WordAp.Selection.TypeText "，僅保留「」等商品，本件商標即有獲准註冊之機會。"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　　" & custtype & "若有意爭取，可刪除「」商品，並據實陳述意見，檢附實際使用之態樣，證明業經長期且廣泛使用，該商標在交易市場上已成為指定商品之識別標識，足資做為消費者辨識商品來源之依據，已具識別性，以爭取之。"
                  Case "2K"      'T核駁前先行通知1202第29條第3項(MXX)及第30條第1項第10款(HXX)
                     g_WordAp.Selection.TypeText "　　智慧局認為以「" & CASENAME & "」作為商標之一部分，指定使用於「" & m_TMGoods & "」等商品，為所指定商品之說明；此外本件商標圖樣中之「" & CASENAME & "」與" & strCP36Kind & "第" & rsA.Fields(2) & "號「" & rsA.Fields(3) & "」商標圖樣相較，……，兩者構圖意匠相彷彿，復指定使用於相同或類似商品或服務，有使一般消費者產生混淆誤認之虞，應屬近似商標。　" & custtype & "可取得經前揭商標所有人同意申請之證明文件，"
                     g_WordAp.Selection.Font.Bold = True
                     g_WordAp.Selection.TypeText "或刪除與前揭商標構成近似之「" & m_TMGoods & "」商品"
                     g_WordAp.Selection.Font.Bold = False
                     g_WordAp.Selection.TypeText "，僅保留「」商品，"
                     g_WordAp.Selection.Font.Bold = True
                     g_WordAp.Selection.TypeText "並聲明本件商標不就「" & CASENAME & "」主張商標權"
                     g_WordAp.Selection.Font.Bold = False
                     g_WordAp.Selection.TypeText "，爭取之；否則以重新設計圖樣，另案申請為宜。"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　　" & custtype & "若已使用該商標，可據實陳述意見，舉證無使消費者混淆誤認之虞之具體事證，並聲明本件商標不就「" & CASENAME & "」主張商標權，爭取之；否則以更改商標圖樣，另案申請為宜。"
                  'Add By Sindy 2012/7/17
                  Case "2L"      'T核駁前先行通知1202第29條第3項(MXX)及第30條第1項第11款(GXX)
                     g_WordAp.Selection.TypeText "　　按「商標相同或近似於他人著名商標或標章，有致相關公眾混淆誤認之虞，或有減損著名商標或標章之識別性或信譽之虞者，不得註冊」為商標法第30條第1項11款所明定。智慧局認為本件商標圖樣上之"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　，有使公眾產生混淆誤認之虞，或有減損著名商標之識別性之虞，此業經智慧局以　　認定在案；此外，以「" & CASENAME & "」作為商標之一部分，指定使用於「" & m_TMGoods & "」等商品，為……之說明，欠缺識別性，且有致商標權範圍產生疑義之虞。本案　" & custtype & "可據實陳述意見，舉證無造成消費者混淆誤認或減損著名商標識別性之虞之具體事證，並"
                     g_WordAp.Selection.Font.Bold = True
                     g_WordAp.Selection.TypeText "同意聲明本件商標不就「" & CASENAME & "」主張商標權"
                     g_WordAp.Selection.Font.Bold = False
                     g_WordAp.Selection.TypeText "，以爭取之；否則以更改商標圖樣，另案申請為宜。"
                  '2012/7/17 End
                  Case "25"      'T核駁前先行通知1202第29條第1項第3款(AXX)
                     g_WordAp.Selection.TypeText "　　智慧局認為商標圖樣上之「" & CASENAME & "」，予消費者的認知，僅為……之描述用語，以之作為商標，指定使用於「" & m_TMGoods & "」等商品，缺乏商標應有的顯著特徵與識別商品或服務來源之功能，並得藉以與他人之商品或服務相區別，欠缺識別性。惟　" & custtype & "若有意爭取，可說明本件商標之設計原由及實際使用態樣，並提出廣告等資料，證明經長期且廣泛使用，該商標在交易市場上已成為指定商品或服務之識別標識，已具識別性，以爭取之；否則以重新設計圖樣，另案申請為宜。"
                  Case "26"      'T核駁前先行通知1202第29條第1項第1款(BXX)
                     g_WordAp.Selection.TypeText "　　智慧局認為商標圖樣上之「" & CASENAME & "」，予消費者的認知，僅為……之描述用語，以之作為商標，指定使用於「" & m_TMGoods & "」等商品，缺乏商標應有的顯著特徵與識別商品或服務來源之功能，並得藉以與他人之商品或服務相區別，欠缺識別性。惟　" & custtype & "若有意爭取，可說明本件商標之設計原由及實際使用態樣，並提出廣告等資料，證明經長期且廣泛使用，該商標在交易市場上已成為指定商品或服務之識別標識，已具識別性，以爭取之；否則以重新設計圖樣，另案申請為宜。"
                  Case "2F"      'T核駁前先行通知1202第29條第1項第2款(CXX)
                     g_WordAp.Selection.TypeText "　　智慧局認為以「" & CASENAME & "」作為商標，指定使用於「" & m_TMGoods & "」等商品，為指定商品或服務之通用標章或名稱，缺乏商標應有的顯著特徵與識別商品或服務來源之功能，並得藉以與他人之商品或服務相區別，欠缺識別性。惟　" & custtype & "若有意爭取，可說明本件商標之設計原由及實際使用態樣，並提出廣告等資料，證明經長期且廣泛使用，該商標在交易市場上已成為指定商品或服務之識別標識，已具識別性，以爭取之；否則以重新設計圖樣，另案申請為宜。"
                  Case "27"      'T核駁前先行通知1202第30條第1項第8款(FXX)
                     g_WordAp.Selection.TypeText "　　智慧局認為以「" & CASENAME & "」作為商標之一部分，指定使用於「" & m_TMGoods & "」商品，有使相關消費者對其表彰商品之性質產生誤認誤信之虞。本案建議　" & custtype & "可減縮「」等商品；或同意刪除圖樣中之「" & CASENAME & "」部分，以爭取本案之獲准註冊。"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　智慧局認為「" & CASENAME & "」係……；以之作為商標，指定使用於「」等商品，有使相關消費者對其表彰商品之產地發生誤認誤信之虞。　" & custtype & "可據實陳述意見，具體說明本件商標之設計原由，並提出無使消費者誤認商品產地之具體事證，爭取之；否則以重新設計圖樣，另案申請為宜。"
                  Case "28"      'T核駁前先行通知1202第30條第1項第11款(GXX)
                     g_WordAp.Selection.TypeText "　　按「商標相同或近似於他人著名商標或標章，有致相關公眾混淆誤認之虞，或有減損著名商標或標章之識別性或信譽之虞者，不得註冊」為商標法第30條第1項11款所明定。智慧局認為本件商標圖樣上之"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　，有使公眾產生混淆誤認之虞，或有減損著名商標之識別性之虞，此業經智慧局以　　認定在案。本案　" & custtype & "可據實陳述意見，舉證無造成消費者混淆誤認或減損著名商標識別性之虞之具體事證，爭取之；否則以更改商標圖樣，另案申請為宜。"
                  Case "29"      'T核駁前先行通知1202第30條第1項第10款(HXX)
                     g_WordAp.Selection.TypeText "　　智慧局認為本件商標圖樣上之「" & CASENAME & "」與" & strCP36Kind & "第" & rsA.Fields(2) & "號「" & rsA.Fields(3) & "」相較，……，復指定使用於相同或類似商品或服務，有使一般消費者產生混淆誤認之虞，應屬近似商標。　" & custtype & "若有意爭取，可據實陳述意見，檢附實際使用之態樣，說明無造成消費者混淆誤認之虞；或"
                     g_WordAp.Selection.Font.Bold = True
                     g_WordAp.Selection.TypeText "刪除與前揭商標構成近似之「" & m_TMGoods & "」商品"
                     g_WordAp.Selection.Font.Bold = False
                     g_WordAp.Selection.TypeText "，僅保留「」商品，以爭取本案之註冊。"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　　" & custtype & "若已使用該商標，可據實陳述意見，舉證無使消費者混淆誤認之虞之具體事證，或取得經前揭商標所有人同意申請之證明文件，爭取之；否則以重新設計圖樣，另案申請為宜。"
                  Case "2A"      'T核駁前先行通知1202第30條第1項第12款(IXX)
                     g_WordAp.Selection.TypeText "　　按「商標相同或近似於他人先使用於同一或類似商品或服務之商標，而申請人因與該他人間具有契約、地緣、業務往來或其他關係，知悉他人商標存在者，意圖仿襲而申請註冊者，不得註冊」，為商標法30條第1項第12款所明定。智慧局認為本件商標之「" & CASENAME & "」與　近似， ，而有違前揭法條之規定。　" & custtype & "可據實陳述意見，舉證無造成混淆誤認之虞；或檢送經「」同意申請之證明文件，爭取之；否則以重新設計圖樣，另案申請為宜。"
                  Case "2B"      'T核駁前先行通知1202第30條第1項第13款(JXX)
                     g_WordAp.Selection.TypeText "　　按「商標有他人之肖像或著名之姓名、藝名、筆名、字號者，不得註冊」為商標法30條第1項第13款所明定。智慧局認為本件商標圖樣係以真實人物照片為構圖，有前揭法條之適用。　" & custtype & "可檢送經該他人同意申請註冊之證明文件，本件商標方有核准註冊之機會。"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "　　按「商標有他人之肖像或著名之姓名、藝名、筆名、字號者，不得註冊」為商標法30條第1項第13款所明定。智慧局認為「 」係　，以「" & CASENAME & "」作為商標之一部分，指定使用於「" & m_TMGoods & "」等商品，有前揭法條之適用。　" & custtype & "可據實陳述意見，具體說明本件商標之設計原由及實際使用態樣，爭取之，否則以重新設計商標圖樣，另案申請為宜。"
                  Case "2I"      'T核駁前先行通知1202第29條第3項(MXX)
                     g_WordAp.Selection.TypeText "　　智慧局認為以「" & CASENAME & "」作為商標之一部分，指定使用於「" & m_TMGoods & "」等商品，為……之說明，欠缺識別性，且有致商標權範圍產生疑義之虞。惟　" & custtype & "若有意爭取，可說明本件商標之設計原由及實際使用態樣，並提出廣告等資料，證明經長期且廣泛使用，該商標在交易市場上已成為指定商品或服務之識別標識，已具識別性，以爭取之；"
                     g_WordAp.Selection.Font.Bold = True
                     g_WordAp.Selection.TypeText "或同意聲明本件商標不就「" & CASENAME & "」主張商標權"
                     g_WordAp.Selection.Font.Bold = False
                     g_WordAp.Selection.TypeText "，本件商標即有獲准註冊之機會。"
                  Case "2Z"      'T核駁前先行通知1202其他條款
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeParagraph
               End Select
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "三、本案本所案號為「" & CaseNo & "」，往後查詢時，請註明本所案號以利處理。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "四、有便煩請　" & custtype & "儘速與本所聯絡，俾決定續行事宜，若尚有其他任何質疑，請隨時告知，本所將竭誠為　" & custtype & "提供服務。"
               '2012/7/11 End
            '2008/9/24 ADD BY SONIA
            Else              '台->大
               StrSQLa = "Select CP07,CP36,NVL(NVL(CP37,CP38),CP39),NVL(NVL(CP40,CP41),CP42) FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                         " AND CP10='1205' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  m_CP07 = Mid(rsA.Fields(0).Value, 1, 4) & "年" & Mid(rsA.Fields(0).Value, 5, 2) & "月" & Mid(rsA.Fields(0).Value, 7, 2) & "日"
               End If
               g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)" & nationname & "商標註冊申請乙案(本所案號" & CaseNo & ")，今接到代理人之通知函，謂本件商標之與大陸商標法規定不合，業經大陸國家知識產權局就該部分駁回，隨函檢附核駁通知書影本乙份，請查照。"
               g_WordAp.Selection.TypeParagraph
               If rsA.Fields(1).Value <> "" Then
                  g_WordAp.Selection.TypeText "二、駁回主要理由：本件商標與" & rsA.Fields(3).Value & "已註冊之第" & rsA.Fields(1).Value & "號「" & rsA.Fields(2).Value & "」商標近似，應駁回在「  」商品之註冊申請。"
               Else
                  g_WordAp.Selection.TypeText "二、駁回理由："
                  g_WordAp.Selection.TypeParagraph
               End If
               g_WordAp.Selection.TypeParagraph
               'Modify By Sindy 2012/4/23
               'g_WordAp.Selection.TypeText "三、本案依法須於" & m_CP07 & "以前向大陸商標評審委員會提出復審，且不得延期。本件商標因  與據以核駁商標   ，指定商品復屬同一或類似，故遭大陸商標局認不得申請註冊，本所以為　" & custtype & "如有意續行辦理，請於隨函所附之委任書上用印，並務請儘速與本所聯繫，以利及時提出相關文書。"
               'modify by sonia 2017/8/16 取消'依法'二字
               g_WordAp.Selection.TypeText "三、本案須於" & m_CP06 & "以前向大陸國家知識產權局會提出復審，且不得延期。本件商標因  與據以核駁商標   ，指定商品復屬同一或類似，故遭大陸國家知識產權局認不得申請註冊，本所以為　" & custtype & "如有意續行辦理，請於隨函所附之委任書上用印，並務請於" & m_DueDate & "以前與本所聯繫，以利及時提出相關文書。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "四、由於期限緊迫，請儘速與本所聯繫，以共商續行事宜。若尚有任何問題，請隨時洽詢，本所竭誠為　" & custtype & "服務。"
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            End If
            '2008/9/24 END
         '2008/9/19 add by sonia
         Case "3"             '勝訴1003,撤銷原處分1402-勝
'2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
'            If m_Combo8 = "3D" Then
'               '2009/3/3 ADD BY SONIA 卷宗性質為申請者再以案件性質判斷casetype3,T-065430
'               If casetype2 = "核駁" Then
'                  StrSQLa = "Select MAX(CP05||CP10) FROM CaseProgress WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
'                            " AND CP10 IN ('101','602','1602','604','1604','606','1606') "
'                  rsA.CursorLocation = adUseClient
'                  rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'                  If rsA.Fields(0) <> "" Then
'                     casetype2 = "註冊"
'                     Select Case Mid(rsA.Fields(0), 9)
'                        Case "602", "1602"
'                           casetype3 = "異議"
'                        Case "604", "1604"
'                           casetype3 = "評定"
'                        Case "606", "1606"
'                           casetype3 = "廢止"
'                     End Select
'                  Else
'                     casetype3 = "　　"    '中間接進來則留空白
'                  End If
'                  If rsA.State <> adStateClosed Then rsA.Close
'                  Set rsA = Nothing
'               End If
'            End If
'            '2009/3/3 END
            '2008/10/17 加入撤銷原處分
            If Combo8.Text = "勝訴　　　　　　1003" Then
               '2009/11/6 modify by sonia T-133711加主管機關
               'StrSQLa = "Select C1.CP08,NVL(NVL(C1.CP40,C1.CP41),C1.CP42),C2.CP10,DECODE(TM10,'000',CPM03,CPM04) FROM CaseProgress C1,CaseProgress C2,CASEPROPERTYMAP,TRADEMARK WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                         " AND C1.CP10='1003' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C' AND C1.CP43=C2.CP09(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C2.CP01=CPM01(+) AND C2.CP10=CPM02(+) "
               StrSQLa = "Select C1.CP08,NVL(NVL(C1.CP40,C1.CP41),C1.CP42),C2.CP10,DECODE(TM10,'000',CPM03,CPM04),CF10 FROM CaseProgress C1,CaseProgress C2,CASEPROPERTYMAP,TRADEMARK,CASEFEE WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                         IIf(m_T727CP43No <> "", " AND C1.CP09='" & m_T727CP43No & "' ", " AND C1.CP10='1003' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C'") & " AND C1.CP43=C2.CP09(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C2.CP01=CPM01(+) AND C2.CP10=CPM02(+) AND C2.CP01=CF01(+) AND '" & pa(10) & "'=CF02(+) AND C2.CP10=CF03(+) "
               '2009/11/6 end
            Else
               '2009/11/6 modify by sonia T-133711加主管機關
               'StrSQLa = "Select C1.CP08,NVL(NVL(C1.CP40,C1.CP41),C1.CP42),C2.CP10,DECODE(TM10,'000',CPM03,CPM04) FROM CaseProgress C1,CaseProgress C2,CASEPROPERTYMAP,TRADEMARK WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                         " AND C1.CP10='1402' AND C1.CP24='1' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C' AND C1.CP43=C2.CP09(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C2.CP01=CPM01(+) AND C2.CP10=CPM02(+) "
               StrSQLa = "Select C1.CP08,NVL(NVL(C1.CP40,C1.CP41),C1.CP42),C2.CP10,DECODE(TM10,'000',CPM03,CPM04),CF10 FROM CaseProgress C1,CaseProgress C2,CASEPROPERTYMAP,TRADEMARK,CASEFEE WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                         IIf(m_T727CP43No <> "", " AND C1.CP09='" & m_T727CP43No & "' AND C1.CP24='1'", " AND C1.CP10='1402' AND C1.CP24='1' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C'") & " AND C1.CP43=C2.CP09(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C2.CP01=CPM01(+) AND C2.CP10=CPM02(+) AND C2.CP01=CF01(+) AND '" & pa(10) & "'=CF02(+) AND C2.CP10=CF03(+) "
               '2009/11/6 end
            End If
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.Fields(2) <> "" Then
               Select Case rsA.Fields(2)
                  Case "602", "1602"
                     'casetype3 = "異議"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                     If pa(10) = "000" Then
                        casepaper = "審定書"
                     Else
                        casepaper = "裁定書"
                     End If
                  Case "604", "1604"
                     If pa(10) = "000" Then
                        'casetype3 = "評定"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                        casepaper = "書"
                    Else
                        'casetype3 = "裁定"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                        '2009/11/5 MODIFY BY SONIA T-133711
                        'casepaper = "裁定書"
                        casepaper = "書"
                     End If
                  'modify by sonia 2019/5/27 +624部分廢止答辯,+1620被部分廢止（理由）
                  Case "606", "1606", "624", "1620"
                     If pa(10) = "000" Then
                        'casetype3 = "廢止"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                        casepaper = "處分書"
                     Else
                        'casetype3 = "撤銷"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                        casepaper = "決定書"     '2017/6/14 modify by sonia 裁定書改決定書 T-147826
                     End If
               End Select
            End If
            Select Case m_Combo8
               Case "31"      'T(異議,評定,廢止)勝訴1003台->台
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & casetype2 & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標" & casetype3 & "乙案(本所案號" & CaseNo & ")，今接獲智慧財產局發給之" & rsA.Fields(0) & "商標" & casetype3 & casepaper & "。茲隨函檢附正本乙份，請查照。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "二、對造如有不服，依法可於收到處分書之次日起三十日內提出訴願。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "三、有關案情本所分析以為："
                  g_WordAp.Selection.TypeParagraph
                  If pa(28) <> "4" Then
                     g_WordAp.Selection.TypeText "　　　經我方說明　　　　　。該等主張獲智慧財產局之採納，遂撤銷系爭商標之註冊。特此轉知。"
                  Else
                     g_WordAp.Selection.TypeText "　　　本案經我方陳明被申請人(公司已廢止登記或已未使用商標)，並舉證調查情事，以系爭商標應已構成停止使用達三年以上之可疑為由，向智慧財產局申請廢止商標之註冊。經該局通知商標權人答辯，未獲回應，乃依法廢止該商標之註冊。特此轉知。"
                  End If
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "四、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               Case "32"      'T(異議,評定,廢止)勝訴1003台->大
                  '2009/11/5 MODIFY BY SONIA T-133711
                  'g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)" & nationname & "商標" & casetype3 & "乙案(本所案號" & caseno & ")，頃接代理人轉來商標局之" & casetype3 & "裁定書，隨函檢附該商標" & casetype3 & "裁定書影本乙紙，請查照。"
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)" & nationname & "商標" & casetype3 & "乙案(本所案號" & CaseNo & ")，頃接代理人轉來" & rsA.Fields(4) & "之" & casetype3 & casepaper & "，隨函檢附該商標" & casetype3 & casepaper & "影本乙紙，請查照。"
                  '2009/11/5 END
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "二、本案經我方極力主張　　　　　，為大陸國家知識產權局所採，而為" & casetype3 & "成立之處分，本案獲得勝訴。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "三、對造若不服，依法可提出復審。若有進一步消息，當立即轉知　" & custtype & "。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "四、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               Case "33"      'T(被異議(理由),被評定(理由),被廢止(理由))勝訴1003台->台
                  g_WordAp.Selection.TypeText "一、" & custtype & "所有註冊" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標前遭" & rsA.Fields(1) & casetype3 & "乙案(本所案號" & CaseNo & ")，頃接智慧財產局之" & casetype3 & casepaper & "，隨函檢附" & rsA.Fields(0) & casetype3 & casepaper & "正本乙紙，請查照。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "二、本案我方雖未提出答辯，惟智慧財產局仍為" & casetype3 & "不成立之處分，本案　" & custtype & "獲得勝訴，特此轉知。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "三、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               Case "34"      'T(被異議(理由),被評定(理由),被廢止(理由))勝訴1003台->大
                  '2009/11/5 MODIFY BY SONIA T-133711
                  'g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)" & nationname & "商標前遭" & rsA.Fields(1) & "提出" & casetype3 & "乙案(本所案號" & caseno & ")，頃接代理人轉來商標局之" & casetype3 & "裁定書，隨函檢附該商標" & casetype3 & "裁定書影本乙紙，請查照。"
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)" & nationname & "商標前遭" & rsA.Fields(1) & "提出" & casetype3 & "乙案(本所案號" & CaseNo & ")，頃接代理人轉來" & rsA.Fields(4) & "之" & casetype3 & casepaper & "，隨函檢附該商標" & casetype3 & casepaper & "影本乙紙，請查照。"
                  '2009/11/5 END
                  g_WordAp.Selection.TypeParagraph
                  'Modified by Lydia 2022/09/28 修改大陸商標異議不成立之定稿
                  'g_WordAp.Selection.TypeText "二、本案我方雖未提出答辯，惟大陸國家知識產權局仍為" & casetype3 & "不成立之處分，本案　" & custtype & "獲得勝訴，若對方未申請復審，依法本件商標將可獲准註冊。俟有進一步消息，當立即通知　" & custtype & "。"
                  g_WordAp.Selection.TypeText "二、本案我方雖未提出答辯，惟大陸國家知識產權局仍為" & casetype3 & "不成立之處分，本案　" & custtype & "獲得勝訴，依法本件商標將可獲准註冊。俟有進一步消息，當立即通知　" & custtype & "。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "三、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               Case "35"      'T(異答,評答,廢答)勝訴1003台->台
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之註冊" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標" & rsA.Fields(3) & "乙案(本所案號" & CaseNo & ")，頃接智慧財產局發給之" & rsA.Fields(0) & "商標" & casetype3 & casepaper & "。茲隨函檢附正本乙份，請查照。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "二、對造如有不服，依法可於收到處分書之次日起三十日內提出訴願。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "三、有關案情本所分析以為："
                  g_WordAp.Selection.TypeParagraph
                  'modify by sonia 2019/5/27 +624部分廢止答辯
                  If rsA.Fields(2) = "606" Or rsA.Fields(2) = "624" Then
                     g_WordAp.Selection.TypeText "　　　本案經我方檢附商標使用資料並提出答辯後，獲智慧財產局採納，認系爭商標未構成應予廢止之情事，遂為申請不成立之處分。特此轉知。"
                  Else
                     g_WordAp.Selection.TypeText "　　　雖對造主張本件商標有違法註冊之情事，惟經我方說明　　　　　。該等理由獲智慧財產局之採納，遂為申請不成立之處分。特此轉知。"
                  End If
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "四、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               Case "36"      'T(異議答辯,評定答辯)勝訴1003台->大  2017/6/14廢止答辯勝訴改至3E(T-147826)
                  '2009/11/5 MODIFY BY SONIA T-133711
                  'g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)" & nationname & "商標" & rsA.Fields(3) & "乙案(本所案號" & caseno & ")，頃接代理人轉來商標局之" & casetype3 & "裁定書，隨函檢附該商標" & casetype3 & "裁定書影本乙紙，請查照。"
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)" & nationname & "商標" & rsA.Fields(3) & "乙案(本所案號" & CaseNo & ")，頃接代理人轉來" & rsA.Fields(4) & "之" & casetype3 & casepaper & "，隨函檢附該商標" & casetype3 & casepaper & "影本乙紙，請查照。"
                  '2009/11/5 END
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "二、本案經我方提出答辯極力爭取，大陸國家知識產權局以　　　　　，而為" & casetype3 & "不成立之處分，本案　" & custtype & "獲得勝訴，若對方未申請復審，依法本件商標將可獲准註冊。俟有進一步消息，當立即通知　" & custtype & "。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "三、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               '2017/6/14 add by sonia T-147826
               Case "3E"      'T(廢止答辯)勝訴1003台->大(T-147826)
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)" & nationname & "商標" & rsA.Fields(3) & "乙案(本所案號" & CaseNo & ")，頃接代理人轉來" & rsA.Fields(4) & "之" & casetype3 & casepaper & "，隨函檢附該商標" & casetype3 & casepaper & "影本乙紙，請查照。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "二、本案經我方提出答辯極力爭取，大陸國家知識產權局採納我方提供之商標使用證據，認無停止使用情形，因而為" & casetype3 & "不成立之決定，本案　" & custtype & "獲得勝訴。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "三、對造若不服，依法可提出復審，若有進一步消息，當立即通知　" & custtype & "。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "四、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               '2017/6/14 end
               Case "37"      'T(申請核駁訴願)勝訴1003台->台
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & casetype2 & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標註冊事件" & rsA.Fields(3) & "乙案(本所案號" & CaseNo & ")，頃接經濟部有關此案之" & rsA.Fields(0) & "決定書乙份，茲隨函檢附正本乙份，敬請查收。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "二、有關案情本所分析如下："
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　　本案經我方說明　　　　　。上述主張果獲經濟部之採納，乃據而將原處分撤銷，智慧財產局將就本案重為處分。特此轉知。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "三、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               Case "38"      'T(訴願)勝訴1003台->大
                  '2016/7/7 MODIFY BY SONIA
                  'g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)" & nationname & "商標註冊申請" & rsA.Fields(3) & "乙案(本所案號" & CaseNo & ")，頃接代理人轉來商標評審委員會之復審終局決定，隨函檢附該商標決定書影本乙紙，請查照。"
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)" & nationname & "商標註冊申請" & rsA.Fields(3) & "乙案(本所案號" & CaseNo & ")，頃接代理人轉來大陸國家知識產權局之復審決定，隨函檢附該商標決定書影本乙紙，請查照。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "二、本案經我方主張　　　　　，為大陸國家知識產權局所採納，本件商標將獲審定公告，俟審定公告後本所當立即轉知。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "三、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               Case "39"      'T(爭議敗訴訴願)勝訴1003台->台
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & casetype2 & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標" & casetype3 & "事件" & rsA.Fields(3) & "乙案(本所案號" & CaseNo & ")，頃接經濟部有關此案之" & rsA.Fields(0) & "決定書乙份，茲隨函檢附正本乙份，敬請查收。"
                  g_WordAp.Selection.TypeParagraph
                  '依對方有無通知參加訴願而決定是否須此段內容, 由操作人員自行刪除
                  g_WordAp.Selection.TypeText "二、對造可於原處分機關重為處分後，決定是否訴願；或因不服本決定，於收到決定書之次日起2個月內逕行提出行政訴訟。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "三、有關案情本所分析如下："
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　　本案經我方於訴願書中說明　　　　　。經濟部於斟酌對造及我方訴願理由後，認本件商標         ，乃據而撤銷原處分，我方獲勝。特此轉知。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "四、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               Case "3B"      'T(爭議行政訴訟)勝訴1003台->台
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & casetype2 & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標" & casetype3 & "事件" & rsA.Fields(3) & "乙案(本所案號" & CaseNo & ")，頃接智慧財產及商業法院有關此案之判決書，茲隨函檢附正本乙份，敬請查收。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "二、對造如不服本案判決，可於判決送達後二十日內以其違背法令為由，向最高行政法院提起上訴。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "三、有關案情本所分析如下："
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　　本案經我方力陳　　　　　。上述主張終獲智慧財產及商業法院之採納，乃據而將原處分及原決均撤銷。若對造未在期限內提出上訴，可望原處分機關將就本案重為處分。特此轉知。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "四、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               Case "3C"      'T(爭議參加訴願)勝訴1003台->台
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & casetype2 & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標" & casetype3 & "事件" & rsA.Fields(3) & "乙案(本所案號" & CaseNo & ")，頃接經濟部有關此案之" & rsA.Fields(0) & "決定書，茲隨函檢附影本乙份，敬請查收。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "二、對造如有不服可於收到決定書之次日起2個月內提出行政訴訟。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "三、有關案情本所分析如下："
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　　經濟部斟酌對造及我方" & rsA.Fields(3) & "之意見，採納我方之理由，認　　　　　，並據而駁回對造之訴願。特此轉知。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "四、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               Case "3D"      'T(爭議行政上訴答辯)勝訴1003台->台
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & casetype2 & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標" & casetype3 & "事件" & rsA.Fields(3) & "乙案(本所案號" & CaseNo & ")，頃接最高行政法院有關此案之判決書，茲隨函檢附正本乙份，敬請查收。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "二、有關案情本所分析如下："
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　　本案經我方就對造之上訴理由提出答辯，說明對造僅就原審對證據之取捨、事實認定等職權行使加以爭執，並未具體陳述法定上訴要件，顯與法不合，且原判決並無違背法令之情事，乃其上訴依法應予駁回。經最高行政法院審理後，採納我方之理由，遂駁回對造之上訴，即我方獲勝，本案至此應告確定。特此轉知。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "三、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
            End Select
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
         '2008/9/19 END
         '2008/10/8 add by sonia
         Case "4"             '敗訴1004,撤銷原處分1402-敗
            'Add By Sindy 2012/4/23
            If m_Combo8 = "4D" Then
               StrSQLa = "Select C1.CP07,C1.CP08,NVL(NVL(C2.CP40,C2.CP41),C2.CP42),C2.CP10,DECODE(TM10,'000',M1.CPM03,M1.CPM04),C3.CP10,C1.CP06 as CP06 FROM CaseProgress C1,CaseProgress C2,CaseProgress C3,CASEPROPERTYMAP M1,TRADEMARK WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                         IIf(m_T727CP43No <> "", " AND C1.CP09='" & m_T727CP43No & "' ", " AND C1.CP10='1004' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C'") & " AND C1.CP43=C2.CP09(+) AND C2.CP43=C3.CP09(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C2.CP01=M1.CPM01(+) AND C2.CP10=M1.CPM02(+) "
            Else
            '2012/4/23 End
               '2008/10/17 加入撤銷原處分
               'modify by sonia 2019/11/19  +"部分勝部分敗　　1006"T-217896
               If Combo8.Text = "敗訴　　　　　　1004" Or Combo8.Text = "部分勝部分敗　　1006" Then
                  '2009/3/2 MODIFY BY SONIA 加下一程序案件性質,T-128044
                  'StrSQLa = "Select C1.CP07,C1.CP08,NVL(NVL(C2.CP40,C2.CP41),C2.CP42),C2.CP10,DECODE(TM10,'000',CPM03,CPM04),C3.CP10 FROM CaseProgress C1,CaseProgress C2,CaseProgress C3,CASEPROPERTYMAP,TRADEMARK WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                            " AND C1.CP10='1004' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C' AND C1.CP43=C2.CP09(+) AND C2.CP43=C3.CP09(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C2.CP01=CPM01(+) AND C2.CP10=CPM02(+) "
                  'Modify By Sindy 2009/08/11 增加本所期限
                  '2010/2/22 MODIFY BY SONIA 加下一程序的主管機關T-132009(裁定敗訴)
                  'StrSQLa = "Select C1.CP07,C1.CP08,NVL(NVL(C2.CP40,C2.CP41),C2.CP42),C2.CP10,DECODE(TM10,'000',M1.CPM03,M1.CPM04),C3.CP10,DECODE(TM10,'000',M2.CPM03,M2.CPM04),C1.CP06 as CP06 FROM CaseProgress C1,CaseProgress C2,CaseProgress C3,CASEPROPERTYMAP M1,CASEPROPERTYMAP M2,TRADEMARK,NEXTPROGRESS WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                            " AND C1.CP10='1004' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C' AND C1.CP43=C2.CP09(+) AND C2.CP43=C3.CP09(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C2.CP01=M1.CPM01(+) AND C2.CP10=M1.CPM02(+) AND C1.CP09=NP01(+) AND NP06 IS NULL AND NP02=M2.CPM01(+) AND NP07=M2.CPM02(+) "
                  'modify by sonia 2017/4/21 台->大"42"承慧給新定稿T-202204,+C1.CP43之F2.CF10
                  'StrSQLa = "Select C1.CP07,C1.CP08,NVL(NVL(C2.CP40,C2.CP41),C2.CP42),C2.CP10,DECODE(TM10,'000',M1.CPM03,M1.CPM04),C3.CP10,DECODE(TM10,'000',M2.CPM03,M2.CPM04),C1.CP06 as CP06,CF10 FROM CaseProgress C1,CaseProgress C2,CaseProgress C3,CASEPROPERTYMAP M1,CASEPROPERTYMAP M2,TRADEMARK,NEXTPROGRESS,CASEFEE WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                            " AND C1.CP10='1004' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C' AND C1.CP43=C2.CP09(+) AND C2.CP43=C3.CP09(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C2.CP01=M1.CPM01(+) AND C2.CP10=M1.CPM02(+) AND C1.CP09=NP01(+) AND NP06 IS NULL AND NP02=M2.CPM01(+) AND NP07=M2.CPM02(+) " & _
                            " AND TM01=CF01(+) AND TM10=CF02(+) AND NP07=CF03 "
                  If m_Combo8 <> "42" Then
                     'modify by sonia 2019/5/22 再加若下一程序已收文T-204490訴願敗訴,下一程序行政訴訟已收文
                     'StrSQLa = "Select C1.CP07,C1.CP08,NVL(NVL(C2.CP40,C2.CP41),C2.CP42),C2.CP10,DECODE(TM10,'000',M1.CPM03,M1.CPM04),C3.CP10,DECODE(TM10,'000',M2.CPM03,M2.CPM04),C1.CP06 as CP06,CF10 FROM CaseProgress C1,CaseProgress C2,CaseProgress C3,CASEPROPERTYMAP M1,CASEPROPERTYMAP M2,TRADEMARK,NEXTPROGRESS,CASEFEE WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                               " AND C1.CP10='1004' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C' AND C1.CP43=C2.CP09(+) AND C2.CP43=C3.CP09(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C2.CP01=M1.CPM01(+) AND C2.CP10=M1.CPM02(+) AND C1.CP09=NP01(+) AND NP06 IS NULL AND NP02=M2.CPM01(+) AND NP07=M2.CPM02(+) " & _
                               " AND TM01=CF01(+) AND TM10=CF02(+) AND NP07=CF03 "
                     'modify by sonia 2019/11/19  C1.CP10='1004'改為C1.CP10 in ('1004','1006'),欄位加C1.CP10
                     StrSQLa = "Select C1.CP07,C1.CP08,NVL(NVL(C2.CP40,C2.CP41),C2.CP42),C2.CP10,DECODE(TM10,'000',M1.CPM03,M1.CPM04),C3.CP10,DECODE(TM10,'000',M2.CPM03,M2.CPM04),C1.CP06 as CP06,CF10,C1.CP10 FROM CaseProgress C1,CaseProgress C2,CaseProgress C3,CaseProgress C4,CASEPROPERTYMAP M1,CASEPROPERTYMAP M2,TRADEMARK,NEXTPROGRESS,CASEFEE WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                               IIf(m_T727CP43No <> "", " AND C1.CP09='" & m_T727CP43No & "' ", " AND C1.CP10 in ('1004','1006') AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C'") & " AND C1.CP43=C2.CP09(+) AND C2.CP43=C3.CP09(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C2.CP01=M1.CPM01(+) AND C2.CP10=M1.CPM02(+) AND C1.CP09=NP01(+) AND NP06 IS NULL AND C1.CP09=C4.CP43(+) AND C4.CP27 IS NULL AND C4.CP57 IS NULL AND NVL(NP02,C4.CP01)=M2.CPM01(+) AND NVL(NP07,C4.CP10)=M2.CPM02(+) " & _
                               " AND TM01=CF01(+) AND TM10=CF02(+) AND NVL(NP07,C4.CP10)=CF03 "
                  Else
                     'modify by sonia 2019/11/22  C1.CP10='1004'改為C1.CP10 in ('1004','1006'),欄位加C1.CP10 T-216042
                     StrSQLa = "Select C1.CP07,C1.CP08,NVL(NVL(C2.CP40,C2.CP41),C2.CP42),C2.CP10,DECODE(TM10,'000',M1.CPM03,M1.CPM04),C3.CP10,DECODE(TM10,'000',M2.CPM03,M2.CPM04),C1.CP06 as CP06,F1.CF10 CF10,F2.CF10 F2CF10,C1.CP10 FROM CaseProgress C1,CaseProgress C2,CaseProgress C3,CASEPROPERTYMAP M1,CASEPROPERTYMAP M2,TRADEMARK,NEXTPROGRESS,CASEFEE F1,CASEFEE F2 WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                               IIf(m_T727CP43No <> "", " AND C1.CP09='" & m_T727CP43No & "' ", " AND C1.CP10 in ('1004','1006') AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C'") & " AND C1.CP43=C2.CP09(+) AND C2.CP43=C3.CP09(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C2.CP01=M1.CPM01(+) AND C2.CP10=M1.CPM02(+) AND C1.CP09=NP01(+) AND NP06 IS NULL AND NP02=M2.CPM01(+) AND NP07=M2.CPM02(+) " & _
                               " AND TM01=F1.CF01(+) AND TM10=F1.CF02(+) AND NP07=F1.CF03 AND TM01=F2.CF01(+) AND TM10=F2.CF02(+) AND C2.CP10=F2.CF03 "
                  End If
                  'end 2017/4/21
               Else
                  'Modify By Sindy 2009/08/11 增加本所期限
                  StrSQLa = "Select C1.CP07,C1.CP08,NVL(NVL(C2.CP40,C2.CP41),C2.CP42),C2.CP10,DECODE(TM10,'000',CPM03,CPM04),C3.CP10,C1.CP06 as CP06 FROM CaseProgress C1,CaseProgress C2,CaseProgress C3,CASEPROPERTYMAP,TRADEMARK WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                            IIf(m_T727CP43No <> "", " AND C1.CP09='" & m_T727CP43No & "' AND C1.CP24='2' ", " AND C1.CP10='1402' AND C1.CP24='2' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C'") & " AND C1.CP43=C2.CP09(+) AND C2.CP43=C3.CP09(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C2.CP01=CPM01(+) AND C2.CP10=CPM02(+) "
               End If
            End If
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               If Not IsNull(rsA.Fields(0).Value) Then
                  'modify by sonia 2019/11/20 判斷年度格式
                  m_CP07 = IIf(bolDateType, Mid(rsA.Fields(0).Value, 1, 4), Mid(rsA.Fields(0).Value, 1, 4) - 1911) & "年" & Mid(rsA.Fields(0).Value, 5, 2) & "月" & Mid(rsA.Fields(0).Value, 7, 2) & "日"
               End If
               'Add By Sindy 2009/08/11
               If Not IsNull(rsA.Fields("CP06").Value) Then
                  'modify by sonia 2019/11/20 判斷年度格式
                  m_CP06 = IIf(bolDateType, Mid(rsA.Fields("CP06").Value, 1, 4), Mid(rsA.Fields("CP06").Value, 1, 4) - 1911) & "年" & Mid(rsA.Fields("CP06").Value, 5, 2) & "月" & Mid(rsA.Fields("CP06").Value, 7, 2) & "日"
                  'Add By Sindy 2012/4/23
                  'Modify By Sindy 2012/8/16 401,402,403的敗訴來函(下一程序：行政訴訟或行政訴訟上訴); 本所期限-3工作天 改為 本所期限-7工作天
                  If rsA.Fields(3) = "401" Or rsA.Fields(3) = "402" Or rsA.Fields(3) = "403" Then
                     'modify by sonia 2017/8/17 台灣案因期限較長維持7工作天,非台灣案改5工作天
                     ''本所期限-7工作天
                     'm_DueDT = CompWorkDay(7, rsA.Fields("CP06").Value, 1)
                     'm_DueDate = Mid(m_DueDT, 1, 4) - 1911 & "年" & Mid(m_DueDT, 5, 2) & "月" & Mid(m_DueDT, 7, 2) & "日"
                     If pa(10) = "000" Then
                        '本所期限-7工作天
                        m_DueDT = CompWorkDay(7, rsA.Fields("CP06").Value, 1)
                        'modify by sonia 2019/11/20 判斷年度格式
                        m_DueDate = IIf(bolDateType, Mid(m_DueDT, 1, 4), Mid(m_DueDT, 1, 4) - 1911) & "年" & Mid(m_DueDT, 5, 2) & "月" & Mid(m_DueDT, 7, 2) & "日"
                     Else
                        '本所期限-5工作天
                        m_DueDT = CompWorkDay(5, rsA.Fields("CP06").Value, 1)
                        m_DueDate = Mid(m_DueDT, 1, 4) & "年" & Mid(m_DueDT, 5, 2) & "月" & Mid(m_DueDT, 7, 2) & "日"
                     End If
                     'end 2017/8/17
                  Else
                  '2012/8/16 End
                     '本所期限-3工作天
                     m_DueDT = CompWorkDay(3, rsA.Fields("CP06").Value, 1)
                     'modify by sonia 2019/11/20 判斷年度格式
                     m_DueDate = IIf(bolDateType, Mid(m_DueDT, 1, 4), Mid(m_DueDT, 1, 4) - 1911) & "年" & Mid(m_DueDT, 5, 2) & "月" & Mid(m_DueDT, 7, 2) & "日"
                  End If
               End If
               '2012/4/23 End
               If rsA.Fields(3) <> "" Then
                  Select Case rsA.Fields(3)
                     Case "602", "1602"
                        casetype2 = "註冊"
                        'casetype3 = "異議"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                        If pa(10) = "000" Then
                           casepaper = "審定書"
                        Else
                           casepaper = "裁定書"
                        End If
                     Case "604", "1604"
                        casetype2 = "註冊"
                        If pa(10) = "000" Then
                           'casetype3 = "評定"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                           casepaper = "書"
                       Else
                           'casetype3 = "裁定"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                           '2009/11/5 MODIFY BY SONIA T-133711
                           'casepaper = "裁定書"
                           casepaper = "書"
                        End If
                     'modify by sonia 2019/5/27 +624部分廢止答辯,+1620被部分廢止（理由）
                     Case "606", "1606", "624", "1620"
                        casetype2 = "註冊"
                        If pa(10) = "000" Then
                           'casetype3 = "廢止"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                           casepaper = "處分書"
                        Else
                           'casetype3 = "撤銷"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                           'modify by sonia 2023/7/24 承慧說裁定書改決定書T-140365
                           casepaper = "決定書"
                        End If
                  End Select
               End If
            End If
            '2010/2/22 MODIFY BY SONIA 所有敗訴都加本所期限
            Select Case m_Combo8
               Case "41"      'T(異議,評定,廢止)敗訴1004台->台
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理對" & casetype2 & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標提出" & casetype3 & "乙案(本所案號" & CaseNo & ")，今接獲智慧財產局發給之" & rsA.Fields(1) & "商標" & casetype3 & casepaper & "。茲隨函檢附正本乙份，請查照。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "二、雖經我方極力主張"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  'Modify By Sindy 2012/4/23
                  'g_WordAp.Selection.TypeText "三、" & custtype & "對上述處分若有不服，應於" & m_CP07 & "以前提出訴願，煩請儘速於" & m_CP06 & "以前與本所聯繫以共商續行事宜。"
                  g_WordAp.Selection.TypeText "三、" & custtype & "對上述處分若有不服，應於" & m_CP06 & "以前提出訴願，煩請於" & m_DueDate & "以前與本所聯繫以共商續行事宜。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "四、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               Case "42"      'T(異議,評定,廢止)敗訴1004台->大
                  'modify by sonia 2017/4/21 承慧給新定稿T-202204
                  'g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理對" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)" & nationname & "商標提出" & casetype3 & "乙案(本所案號" & CaseNo & ")，頃接代理人轉來商標局之" & casetype3 & casepaper & "，隨函檢附該商標" & casetype3 & casepaper & "影本乙紙，請查照。"
                  'g_WordAp.Selection.TypeParagraph
                  'g_WordAp.Selection.TypeText "二、我方雖主張　　　　　；商標局以　　　　　而為駁回本件" & casetype3 & "案之處分。"
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理對" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)" & nationname & "商標提出" & rsA.Fields(4) & "乙案(本所案號" & CaseNo & ")，頃接代理人轉來" & rsA.Fields("F2CF10") & "之" & rsA.Fields(4) & "請求" & casepaper & "，隨函檢附該商標" & casepaper & "影本乙紙，請查照。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "二、我方雖主張　　　　　；" & rsA.Fields("F2CF10") & "以　　　　　而為" & rsA.Fields(4) & "不成立之裁定。"
                  'end 2017/4/21
                  g_WordAp.Selection.TypeParagraph
                  '2010/2/22 MODIFY BY SONIA 大陸裁定下一程序為403起訴,其主管機關與異議,廢止不同T-132009
                  'g_WordAp.Selection.TypeText "三、" & custtype & "對於該處分若有不服，須於" & m_CP07 & "以前向商標評審委員會提出" & rsA.Fields(6) & "。　" & custtype & "若有意提出" & rsA.Fields(6) & "，則必須再加強蒐集　　　　　以反駁商標局之理由。"
                  'Modify By Sindy 2012/4/23
                  'g_WordAp.Selection.TypeText "三、" & custtype & "對於該處分若有不服，須於" & m_CP07 & "以前向" & rsA.Fields("CF10") & "提出" & rsA.Fields(6) & "。　" & custtype & "若有意提出" & rsA.Fields(6) & "，則必須再加強蒐集　　　　　以反駁商標局之理由。"
                  g_WordAp.Selection.TypeText "三、" & custtype & "對於該處分若有不服，須於" & m_CP06 & "以前向" & rsA.Fields("CF10") & "提出" & rsA.Fields(6) & "。　" & custtype & "若有意提出" & rsA.Fields(6) & "，則必須再加強蒐集　　　　　以反駁" & rsA.Fields("F2CF10") & "之理由。"
                  '2010/2/22 END
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "四、由於期限緊迫，請儘速於" & m_DueDate & "以前與本所聯繫，以共商續行事宜。若尚有任何問題，請隨時洽詢，本所竭誠為　" & custtype & "服務。"
               '2015/7/15 ADD BY SONIA 大陸異議敗訴無期限,改新格式 T-183743
               Case "42A"      'T(異議)敗訴1004台->大
                  m_CP06 = CompDate(1, 3, strSrvDate(1))
                  m_CP06 = Mid(m_CP06, 1, 4) & "年" & Mid(m_CP06, 5, 2) & "月" & Mid(m_CP06, 7, 2) & "日"
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理對" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)" & nationname & "商標提出" & casetype3 & "乙案(本所案號" & CaseNo & ")，頃接代理人轉來大陸國家知識產權局之" & casetype3 & casepaper & "，隨函檢附該商標" & casetype3 & casepaper & "影本乙紙，請查照。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "二、我方雖主張　　　　　。惟大陸國家知識產權局以　　　　　仍核准對造商標之註冊。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "三、" & custtype & "對於該處分若有不服，可向大陸國家知識產權局會提出無效宣告。若有意續行該一程序，除就原理由再強調外，更宜積極蒐集　　　　　　　，以利證明系爭商標之註冊確應予撤銷。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "四、請於" & m_CP06 & "以前與本所聯繫，以共商續行事宜。若尚有任何問題，請隨時洽詢，本所竭誠為　" & custtype & "服務。"
               '2015/7/15 END
               Case "43"      'T(被異議(理由),被評定(理由),被廢止(理由))敗訴1004台->台
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & casetype2 & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標前遭" & rsA.Fields(2) & "提出" & casetype3 & "乙案(本所案號" & CaseNo & ")，今接獲智慧財產局發給之" & rsA.Fields(1) & "商標" & casetype3 & casepaper & "。茲隨函檢附正本乙份，請查照。"
                  g_WordAp.Selection.TypeParagraph
                  'Modify By Sindy 2012/4/23
                  'g_WordAp.Selection.TypeText "二、依法我方須於" & m_CP07 & "以前提出訴願，否則本件商標之註冊將被撤銷確定。事關　" & custtype & "權益，煩請儘速於" & m_CP06 & "以前與本所聯繫，以共商續行事宜。"
                  'modify by sonia 2019/11/19 T-217896部分勝部分敗1006
                  g_WordAp.Selection.TypeText "二、依法我方須於" & m_CP06 & "以前提出訴願，否則本件商標" & IIf(rsA.Fields(9).Value = "1006", "(於部分商品／服務)", "") & "之註冊將被撤銷確定。事關　" & custtype & "權益，煩請於" & m_DueDate & "以前與本所聯繫，以共商續行事宜。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "三、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               Case "44"      'T(被異議(理由),被評定(理由),被廢止(理由))敗訴1004台->大  撤銷
                  'Modify By Sindy 2018/6/26 此段程式中,凡有rsA.Fields增加IIf(rsA.RecordCount > 0,xxx," ")判斷 ex:T-069020/CA7036866
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)" & nationname & "商標前遭" & IIf(rsA.RecordCount > 0, rsA.Fields(2), " ") & "提出" & casetype3 & "乙案(本所案號" & CaseNo & ")，頃接代理人轉來大陸國家知識產權局之" & casetype3 & casepaper & "，隨函檢附該商標" & casetype3 & casepaper & "影本乙紙，請查照。"
                  g_WordAp.Selection.TypeParagraph
                  'modify by sonia 2017/4/20 撤銷未提出答辯敗訴 T-176068
                  'g_WordAp.Selection.TypeText "二、本案由於我方未提出答辯，商標局以　　　　　，而為本件商標不予註冊之處分。"
                  If casetype3 = "撤銷" Then
                     g_WordAp.Selection.TypeText "二、本案由於我方未提出答辯，大陸國家知識產權局以　" & custtype & "未提交本件商標在指定商品之使用證據，乃撤銷本件商標之註冊，原註冊證作廢。"
                  Else
                     g_WordAp.Selection.TypeText "二、本案由於我方未提出答辯，大陸國家知識產權局依職權審查，以　　　　　，而為本件商標不予註冊之處分。"
                  End If
                  'end 2017/4/20
                  g_WordAp.Selection.TypeParagraph
                  '2010/2/22 MODIFY BY SONIA 大陸裁定下一程序為403起訴,其主管機關與異議,廢止不同T-132009
                  'g_WordAp.Selection.TypeText "三、" & custtype & "對於該處分若有不服，可於" & m_CP07 & "以前向商標評審委員會提出復審。　" & custtype & "若有意提出復審，請儘速與本所聯繫，以共商續行事宜。"
                  'Modify By Sindy 2012/4/23
                  'g_WordAp.Selection.TypeText "三、" & custtype & "對於該處分若有不服，可於" & m_CP07 & "以前向" & rsA.Fields("CF10") & "提出" & rsA.Fields(6) & "。　" & custtype & "若有意提出" & rsA.Fields(6) & "，請儘速與本所聯繫，以共商續行事宜。"
                  g_WordAp.Selection.TypeText "三、" & custtype & "對於該處分若有不服，可於" & m_CP06 & "以前向" & IIf(rsA.RecordCount > 0, rsA.Fields("CF10"), " ") & "提出" & IIf(rsA.RecordCount > 0, rsA.Fields(6), " ") & "。　" & custtype & "若有意提出" & IIf(rsA.RecordCount > 0, rsA.Fields(6), " ") & "，請於" & m_DueDate & "以前與本所聯繫，以共商續行事宜。"
                  g_WordAp.Selection.TypeParagraph
                  'modify by sonia 2015/9/22 m_CP06已於第三段提到,不必再提
                  'g_WordAp.Selection.TypeText "四、由於期限緊迫，請儘速於" & m_CP06 & "以前與本所聯繫，以共商續行事宜。若尚有任何問題，請隨時洽詢，本所竭誠為　" & custtype & "服務。"
                  g_WordAp.Selection.TypeText "四、若尚有任何問題，請隨時洽詢，本所竭誠為　" & custtype & "服務。"
                  'end 2015
               Case "45"      'T(異答,評答,廢答)敗訴1004台->台
                  If rsA.Fields(3) = "602" Then
                     g_WordAp.Selection.TypeText "一、" & custtype & "委託本所辦理" & casetype2 & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標(本所案號" & CaseNo & ")，前遭" & rsA.Fields(2) & "提出" & casetype3 & "，今接獲智慧財產局發給之" & rsA.Fields(1) & "商標" & casetype3 & casepaper & "。茲隨函檢附正本乙份，請查照。"
                  Else
                     g_WordAp.Selection.TypeText "一、" & custtype & "委託本所辦理" & casetype2 & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標(本所案號" & CaseNo & ")，前遭" & rsA.Fields(2) & "申請" & casetype3 & "，今接獲智慧財產局發給之" & rsA.Fields(1) & "商標" & casetype3 & casepaper & "。茲隨函檢附正本乙份，請查照。"
                  End If
                  g_WordAp.Selection.TypeParagraph
                  'Modify By Sindy 2012/4/23
                  'g_WordAp.Selection.TypeText "二、本案依法我方可於" & m_CP07 & "以前提出訴願。"
                  g_WordAp.Selection.TypeText "二、本案依法我方可於" & m_CP06 & "以前提出訴願。　" & custtype & "若有意續行，請於" & m_DueDate & "以前與本所聯繫，以共商相關事宜。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "三、有關案情本所分析以為："
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "四、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               Case "46"      'T(異議答辯,評定答辯,廢止答辯)敗訴1004台->大
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)" & nationname & "商標乙案(本所案號" & CaseNo & ")，前遭" & rsA.Fields(2) & "提出" & casetype3 & "，頃接代理人轉來大陸國家知識產權局之" & casetype3 & casepaper & "，隨函檢附該商標" & casetype3 & casepaper & "影本乙紙，請查照。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "二、我方雖主張　　　　　；惟不為大陸國家知識產權局所採，其以　　　　　而為撤銷本件商標之處分。"
                  g_WordAp.Selection.TypeParagraph
                  '2010/2/22 MODIFY BY SONIA 大陸裁定下一程序為403起訴,其主管機關與異議,廢止不同T-132009
                  'g_WordAp.Selection.TypeText "三、" & custtype & "對於該處分若有不服，可於" & m_CP07 & "以前向商標評審委員會提出復審。　" & custtype & "若有意提出復審，請儘速與本所聯繫，以共商續行事宜。"
                  'Modify By Sindy 2012/4/23
                  'g_WordAp.Selection.TypeText "三、" & custtype & "對於該處分若有不服，可於" & m_CP07 & "以前向" & rsA.Fields("CF10") & "提出" & rsA.Fields(6) & "。　" & custtype & "若有意提出" & rsA.Fields(6) & "，請儘速與本所聯繫，以共商續行事宜。"
                  g_WordAp.Selection.TypeText "三、" & custtype & "對於該處分若有不服，可於" & m_CP06 & "以前向" & rsA.Fields("CF10") & "提出" & rsA.Fields(6) & "。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "四、由於期限緊迫，請儘速於" & m_DueDate & "以前與本所聯繫，以共商續行事宜。若尚有任何問題，請隨時洽詢，本所竭誠為　" & custtype & "服務。"
               Case "47"      'T(申請核駁訴願,爭議敗訴訴願)敗訴1004台->台
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & casetype2 & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標" & casetype3 & "事件" & rsA.Fields(4) & "乙案(本所案號" & CaseNo & ")，頃接經濟部有關此案之" & rsA.Fields(1) & "決定書乙份，茲隨函檢附正本乙份，敬請查收。"
                  g_WordAp.Selection.TypeParagraph
                  'Modify By Sindy 2012/4/23
                  'g_WordAp.Selection.TypeText "二、本案依法我方可於" & m_CP07 & "以前向智慧財產及商業法院提起行政訴訟。"
                  g_WordAp.Selection.TypeText "二、本案依法我方可於" & m_CP06 & "以前向智慧財產及商業法院提起行政訴訟。　" & custtype & "若有意續行，請於" & m_DueDate & "以前與本所聯繫，以共商相關事宜。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "三、有關案情本所分析以為："
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　　雖經我方說明　　　　　。惟經濟部仍以　　　　　，乃駁回我方之" & rsA.Fields(4) & "。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "四、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               Case "48"      'T(訴願)敗訴1004台->大
                  '2016/7/7 MODIFY BY SONIA
                  'g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)" & nationname & "商標註冊申請" & rsA.Fields(4) & "乙案(本所案號" & CaseNo & ")，頃接代理人轉來商標評審委員會之復審終局決定，隨函檢附該商標決定書影本乙紙，請查照。"
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)" & nationname & "商標註冊申請" & rsA.Fields(4) & "乙案(本所案號" & CaseNo & ")，頃接代理人轉來大陸國家知識產權局之復審決定，隨函檢附該商標決定書影本乙紙，請查照。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "二、本案雖經我方主張　　　　　，惟不為大陸國家知識產權局所採，　　　　　，而駁回" & rsA.Fields(4) & "。"
                  g_WordAp.Selection.TypeParagraph
                  'Add By Sindy 2009/08/11
   '2009/12/1 CANCEL BY SONIA T-150442
   '                  g_WordAp.Selection.TypeText "三、由於大陸現" & rsA.Fields(4) & "的審查速度加快，無法等到撤銷案的結果，　" & custtype & "不妨重新提出商標註冊申請。"
   '                  g_WordAp.Selection.TypeParagraph
   '2009/12/1 END
                  'Modify By Sindy 2012/4/23
                  'g_WordAp.Selection.TypeText "三、" & custtype & "對於該處分若有不服，可於" & m_CP07 & "以前向北京市第一中級人民法院起訴。"
                  'MODIFY BY SONIA 2014/12/27 主管機關改抓CASEFEE之起訴程序的主管機關,並刪除後面續行字樣 T-181096
                  'g_WordAp.Selection.TypeText "三、" & custtype & "對於該處分若有不服，可於" & m_CP06 & "以前向北京市第一中級人民法院起訴。　" & custtype & "若有意續行，請於" & m_DueDate & "以前與本所聯繫，以共商相關事宜。"
                  g_WordAp.Selection.TypeText "三、" & custtype & "對於該處分若有不服，可於" & m_CP06 & "以前向" & rsA.Fields("CF10") & "起訴。"
                  'END 2014/12/27
                  g_WordAp.Selection.TypeParagraph
                  'Modify By Sindy 2009/08/11
                  'g_WordAp.Selection.TypeText "四、由於期限緊迫，請儘速與本所聯繫，以共商續行事宜。若尚有任何問題，請隨時洽詢，本所竭誠為　" & custtype & "服務。"
                  'Modify By Sindy 2012/5/30
                  'g_WordAp.Selection.TypeText "四、由於期限緊迫，請儘速於" & m_CP06 & "以前與本所聯繫，以共商續行事宜。若尚有任何問題，請隨時洽詢，本所竭誠為　" & custtype & "服務。"
                  g_WordAp.Selection.TypeText "四、由於期限緊迫，請儘速於" & m_DueDate & "以前與本所聯繫，以共商續行事宜。若尚有任何問題，請隨時洽詢，本所竭誠為　" & custtype & "服務。"
               Case "4B"      'T(爭議行政訴訟，爭議參加訴訟)敗訴1004台->台
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & casetype2 & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標" & casetype3 & "事件提起" & rsA.Fields(4) & "乙案(本所案號" & CaseNo & ")，頃接智慧財產及商業法院有關此案之" & rsA.Fields(1) & "判決書，茲隨函檢附正本乙份，敬請查收。"
                  g_WordAp.Selection.TypeParagraph
                  'Modify By Sindy 2012/4/23
                  'g_WordAp.Selection.TypeText "二、" & custtype & "如認原判決違背法令，欲委託本所提起上訴，請於" & m_CP07 & "以前與本所連繫，以免延誤法定上訴期間。"
                  g_WordAp.Selection.TypeText "二、" & custtype & "如認原判決違背法令，欲委託本所提起上訴，請於" & m_DueDate & "以前與本所連繫，以免延誤法定上訴期間。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "三、有關案情本所分析以為："
                  g_WordAp.Selection.TypeParagraph
                  Select Case rsA.Fields(3)
                     Case "403"   '行政訴訟
                        g_WordAp.Selection.TypeText "　　　雖經我方力陳　　　　　。我方對此一判決如有不服，雖可提出上訴，但本案欲以判決違背法令加以爭執，有一定之困難，請慎重考量。"
                     Case ""
                        g_WordAp.Selection.TypeText "　　　雖經我方參加訴訟，說明　　　　　。"
                        g_WordAp.Selection.TypeParagraph
                        g_WordAp.Selection.TypeText "　　　" & custtype & "可慎重考量如何續行：是待被告機關重為處分後，再視結果決定行止；或斟酌此一判決有無違背法令之情事，逕提出上訴加以爭執。"
                  End Select
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "四、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               Case "4C"      'T(通知參加訴願(未參加))敗訴1004台->台
                  Select Case rsA.Fields(5)
                     Case "602", "1602"
                        g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之註冊" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標(本所案號" & CaseNo & ")，前遭異議並獲異議不成立之審定乙案，頃接經濟部所發給有關本案之" & rsA.Fields(1) & "訴願決定書。茲隨函檢附正本乙份，請查照。"
                     Case "604", "1604"
                        g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之註冊" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標(本所案號" & CaseNo & ")，前遭申請評定並獲申請不成立之評決乙案，頃接經濟部所發給有關本案之" & rsA.Fields(1) & "訴願決定書。茲隨函檢附正本乙份，請查照。"
                     'modify by sonia 2019/5/27 +624部分廢止答辯,+1620被部分廢止（理由）
                     Case "606", "1606", "624", "1620"
                        g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之註冊" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標(本所案號" & CaseNo & ")，前遭申請廢止並獲申請不成立之處分乙案，頃接經濟部所發給有關本案之" & rsA.Fields(1) & "訴願決定書。茲隨函檢附正本乙份，請查照。"
                  End Select
                  g_WordAp.Selection.TypeParagraph
                  'Modify By Sindy 2012/4/23
                  'g_WordAp.Selection.TypeText "二、本案依法我方可於" & m_CP07 & "前逕行向智慧財產及商業法院提起行政訴訟，亦可待智慧財產局重新處分後，再決定如何續行。"
                  g_WordAp.Selection.TypeText "二、本案依法我方可於" & m_CP06 & "前逕行向智慧財產及商業法院提起行政訴訟，亦可待智慧財產局重新處分後，再決定如何續行。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "三、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               'Add By Sindy 2012/4/23
               Case "4D"      'T行政訴訟上訴敗訴408  台->台
                  'Modified by Lydia 2020/10/15 委由=>委託
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之註冊" & AppNo & "「" & CASENAME & "」" & casetype4 & "商標異議事件行政訴訟上訴乙案(本所案號" & CaseNo & ")，頃接最高行政法院有關之裁判書乙份，茲隨函附寄正本請查照。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "二、有關案情本所分析以為："
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　　　雖經我方說明　　。惟最高行政法院認為各該主張屬　　，認不符合提起上訴之要件，乃為上訴駁回之裁定。本案之爭訟至此已告一段落，特此轉知。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "三、其他任何質疑，請隨時賜教，本所將竭誠為　" & custtype & "服務。"
               '2012/4/23 End
            End Select
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
         '2008/10/8 end
         'Modify By Sindy 2011/2/25
         Case Else
            Call WordChinese_sub2(g_WordAp)
      End Select
      '2008/6/25 END
End Sub

Private Sub WordChinese_sub2(g_WordAp As Word.Application)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim m_CP07 As String    '法定期限/中民
Dim m_CP06 As String    '本所期限/中民
Dim m_TIME As String    '開庭時間2008/11/12 ADD BY SONIA
Dim m_PLACE As String   '開庭地點2008/11/12 ADD BY SONIA
Dim i As String         '2008/11/12 ADD BY SONIA
Dim m_DueDT As String
Dim m_DueDate As String
Dim m_CP64 As String    '部分核駁商品2009/6/29 add by sonia
Dim m_GoodsName As String
Dim m_strCPM03 As String 'Add By Sindy 2015/1/14
Dim casecopy As String   '1602,1604來函文書  add by sonia  2017/8/22
   
'*******2019/11/19 統一台灣案用民國年月日,但外至台改用西元年月日,非台灣案用西元年月日
   Select Case Left(m_Combo8, 1)
      '2008/9/24 add by sonia
      Case "5"
         Select Case m_Combo8
            Case "51"      '核駁1002台->大
               'Modify By Sindy 2009/08/11 增加本所期限
               'Modified by Lydia 2017/04/27 抓複審費用
               'StrSQLa = "Select CP07,CP36,NVL(NVL(CP37,CP38),CP39),NVL(NVL(CP40,CP41),CP42),CP06 FROM CaseProgress WHERE " & ChgCaseProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                         " AND CP10='1002' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' "
               StrSQLa = "Select CP07,CP36,NVL(NVL(CP37,CP38),CP39),NVL(NVL(CP40,CP41),CP42),CP06,(NVL(CF08,0) + NVL(CF13,0) * 1000) CFEE FROM CaseProgress,CASEFEE WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                         " AND CP10='1002' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' AND CP01=CF01(+) AND CF02='" & pa(10) & "' AND CF03='401' "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  m_CP07 = Mid(rsA.Fields(0).Value, 1, 4) & "年" & Mid(rsA.Fields(0).Value, 5, 2) & "月" & Mid(rsA.Fields(0).Value, 7, 2) & "日"
                  'Add By Sindy 2009/08/11
                  m_CP06 = Mid(rsA.Fields("CP06").Value, 1, 4) & "年" & Mid(rsA.Fields("CP06").Value, 5, 2) & "月" & Mid(rsA.Fields("CP06").Value, 7, 2) & "日"
                  'Add By Sindy 2012/4/23
                  '本所期限-3工作天
                  m_DueDT = CompWorkDay(3, rsA.Fields("CP06").Value, 1)
                  m_DueDate = Mid(m_DueDT, 1, 4) & "年" & Mid(m_DueDT, 5, 2) & "月" & Mid(m_DueDT, 7, 2) & "日"
                  '2012/4/23 End
               End If
               g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)" & nationname & "商標註冊申請乙案(本所案號" & CaseNo & ")，今接到代理人之通知函，謂本件商標與大陸商標法規定不合，業經大陸國家知識產權局駁回，隨函檢附核駁通知書影本乙份，請查照。"
               g_WordAp.Selection.TypeParagraph
               If IsNull(rsA.Fields(1).Value) Then   '無對造
                  g_WordAp.Selection.TypeText "二、駁回理由："
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  'Add By Sindy 2009/08/11
                  'g_WordAp.Selection.TypeText "三、本案依法須於" & m_CP07 & "以前向大陸商標評審委員會申請復審，且不得延期。由於　　　。本所以為　" & custtype & "如有意續行辦理，請於隨函所附之委任書上用印，並務請儘速與本所聯繫，以利及時提出相關文書。"
                  'Modify By Sindy 2012/4/23
                  'g_WordAp.Selection.TypeText "三、本案依法須於" & m_CP07 & "以前向大陸商標評審委員會申請復審，且不得延期。由於　　　。本所以為　。"
                  'modify by sonia 2017/8/16 取消'依法'二字
                  'modify by sonia 2019/5/24 所有的商標評審委員會改為國家知識產權局
                  g_WordAp.Selection.TypeText "三、本案須於" & m_CP06 & "以前向大陸國家知識產權局申請復審，且不得延期。由於　　　。本所以為　。"
               Else   '有對造
                  g_WordAp.Selection.TypeText "二、駁回主要理由：本件商標與" & rsA.Fields(3).Value & "已註冊之第" & rsA.Fields(1).Value & "號「" & rsA.Fields(2).Value & "」商標近似，復指定類似商品，故應駁回註冊申請。"
                  g_WordAp.Selection.TypeParagraph
                  'Add By Sindy 2009/08/11
                  'g_WordAp.Selection.TypeText "三、本案依法須於" & m_CP07 & "以前向大陸商標評審委員會申請復審，且不得延期。本件商標因   與據以核駁商標   ，指定商品復屬同一或類似，故遭大陸商標局認不得申請註冊。本所以為　" & custtype & "如有意續行辦理，請於隨函所附之委任書上用印，並務請儘速與本所聯繫，以利及時提出相關文書。"
                  'Modify By Sindy 2012/4/23
                  'g_WordAp.Selection.TypeText "三、本案依法須於" & m_CP07 & "以前向大陸商標評審委員會申請復審，且不得延期。本件商標因   與據以核駁商標   ，指定商品復屬同一或類似，故遭大陸商標局認不得申請註冊。本所以為　。"
                  'modify by sonia 2017/8/16 取消'依法'二字
                  'modify by sonia 2019/5/24 所有的商標局會改為國家知識產權局
                  g_WordAp.Selection.TypeText "三、本案須於" & m_CP06 & "以前向大陸國家知識產權局申請復審，且不得延期。本件商標因   與據以核駁商標   ，指定商品復屬同一或類似，故遭大陸國家知識產權局認不得申請註冊。本所以為　。"
               End If
               g_WordAp.Selection.TypeParagraph
               'Add By Sindy 2009/08/11
               'Modify By Sindy 2010/7/19 刪除[及申請書]字樣
               'g_WordAp.Selection.TypeText "四、" & custtype & "如有意續行辦理，請於隨函所附之委任書及申請書上用印，並務請儘速與本所聯繫，以利及時提出相關文書。"
               'modify by sonia 2014/10/27
               'g_WordAp.Selection.TypeText "四、" & custtype & "如有意續行辦理，請於隨函所附之委任書上用印，並務請儘速與本所聯繫，以利及時提出相關文書。"
               '2015/8/12 modify by sonia 加入固定費用
               'g_WordAp.Selection.TypeText "四、" & custtype & "如有意續行辦理，請於隨函所附之委任書上蓋公司章及代表人章，與已用印之公司登記資料影本一併提供予本所，並務請儘速聯繫，以利及時提出相關文書。"
               'modify by sonia 2016/9/19 28,000->29,000
               'Modified by Lydia 2017/04/27 費用改抓設定CFEE
               'g_WordAp.Selection.TypeText "四、" & custtype & "如有意續行辦理，請於隨函所附之委任書上蓋公司章及代表人章，與已用印之公司登記資料影本一併提供予本所，並務請儘速聯繫，以利及時提出相關文書，此一程序的費用為29,000元。"
               g_WordAp.Selection.TypeText "四、" & custtype & "如有意續行辦理，請於隨函所附之委任書上蓋公司章及代表人章，與已用印之公司登記資料影本一併提供予本所，並務請儘速聯繫，以利及時提出相關文書，此一程序的費用為" & Format("" & rsA.Fields("CFEE"), "###,##0") & "元。"
               'end 2014/10/27
               g_WordAp.Selection.TypeParagraph
               'Modify By Sindy 2009/08/11
               'g_WordAp.Selection.TypeText "四、由於期限緊迫，請儘速與本所聯繫，以共商續行事宜。若尚有任何問題，請隨時洽詢，本所竭誠為　" & custtype & "服務。"
               'Modify By Sindy 2012/4/23
               'g_WordAp.Selection.TypeText "五、由於期限緊迫，請於" & m_CP06 & "之前儘速與本所聯繫，若尚有任何問題，請隨時洽詢，本所竭誠為　" & custtype & "服務。"
               g_WordAp.Selection.TypeText "五、由於期限緊迫，請於" & m_DueDate & "之前與本所聯繫，若尚有任何問題，請隨時洽詢，本所竭誠為　" & custtype & "服務。"
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            Case "52"      '部分核駁1205
               '2009/6/29 modify by sonia 加部分核駁商品
               'Modify By Sindy 2009/08/11 增加本所期限
               'Modified by Lydia 2017/04/27 抓複審費用
               'StrSQLa = "Select CP07,CP36,NVL(NVL(CP37,CP38),CP39),NVL(NVL(CP40,CP41),CP42),CP64,CP06 FROM CaseProgress WHERE " & ChgCaseProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                         " AND CP10='1205' AND CP27 IS NULL AND CP57 IS NUL L AND CP09>'C' "
               StrSQLa = "Select CP07,CP36,NVL(NVL(CP37,CP38),CP39),NVL(NVL(CP40,CP41),CP42),CP64,CP06,(NVL(CF08,0) + NVL(CF13,0) * 1000) CFEE FROM CaseProgress,CASEFEE WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                         " AND CP10='1205' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' AND CP01=CF01(+) AND CF02='" & pa(10) & "' AND CF03='401' "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  m_CP07 = Mid(rsA.Fields(0).Value, 1, 4) & "年" & Mid(rsA.Fields(0).Value, 5, 2) & "月" & Mid(rsA.Fields(0).Value, 7, 2) & "日"
                  m_CP64 = "" & Mid(rsA.Fields(4).Value, 8)
                  'Add By Sindy 2009/08/11
                  m_CP06 = Mid(rsA.Fields("CP06").Value, 1, 4) & "年" & Mid(rsA.Fields("CP06").Value, 5, 2) & "月" & Mid(rsA.Fields("CP06").Value, 7, 2) & "日"
                  'Add By Sindy 2012/4/23
                  '本所期限-3工作天
                  m_DueDT = CompWorkDay(3, rsA.Fields("CP06").Value, 1)
                  m_DueDate = Mid(m_DueDT, 1, 4) & "年" & Mid(m_DueDT, 5, 2) & "月" & Mid(m_DueDT, 7, 2) & "日"
                  '2012/4/23 End
               End If
               m_GoodsName = GetGoodsName(pa(1), pa(2), pa(3), pa(4)) 'Add By Sindy 2009/07/02
               g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)" & nationname & "商標註冊申請乙案(本所案號" & CaseNo & ")，今接到代理人之通知函，謂本件商標之部分申請與大陸商標法規定不合，業經大陸國家知識產權局就該部分駁回，隨函檢附核駁通知書影本乙份，請查照。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "二、駁回主要理由：本件商標與" & rsA.Fields(3).Value & "已註冊之第" & rsA.Fields(1).Value & "號「" & rsA.Fields(2).Value & "」商標近似，應駁回在「" & m_CP64 & "」商品之註冊申請。"
               g_WordAp.Selection.TypeParagraph
'2014/11/27 modify by sonia 整段修改 T-187949
'               'Modify By Sindy 2009/08/11
'               'g_WordAp.Selection.TypeText "三、本案依法須於" & m_CP07 & "以前向大陸商標評審委員會申請復審，且不得延期。由於本件商標　與引証商標    ，故經大陸商標局認定近似，類似部分之商品不得註冊。本所以為　" & custtype & "如有意續行辦理，請於隨函所附之委任書上用印，並務請儘速與本所聯繫。復審期間，初步審定暫不公告，而如屆期未提出復審，本件商標將在「" & m_GoodsName & "」商品獲准審定。"
'               'Modify By Sindy 2012/4/23
'               'g_WordAp.Selection.TypeText "三、本案依法須於" & m_CP07 & "以前向大陸商標評審委員會申請復審，且不得延期。由於本件商標與引証商標    ，故經大陸商標局認定近似，類似部分之商品不得註冊。本所以為　 ，應不易爭取。"
'               g_WordAp.Selection.TypeText "三、本案依法須於" & m_CP06 & "以前向大陸商標評審委員會申請復審，且不得延期。由於本件商標與引証商標    ，故經大陸商標局認定近似，類似部分之商品不得註冊。本所以為　 ，應不易爭取。"
'               g_WordAp.Selection.TypeParagraph
'               'Add By Sindy 2009/08/11
'               'Modify By Sindy 2010/7/19 刪除[及申請書]字樣
'               'g_WordAp.Selection.TypeText "四、" & custtype & "如有意續行辦理，請於隨函所附之委任書及申請書上用印，並務請儘速與本所聯繫。"
'               'modify by sonia 2014/10/27
'               'g_WordAp.Selection.TypeText "四、" & custtype & "如有意續行辦理，請於隨函所附之委任書上用印，並務請儘速與本所聯繫。"
'               If custtype = "台端" Then
'                  g_WordAp.Selection.TypeText "四、" & custtype & "如有意續行辦理，請於隨函所附之委任書上蓋章，與已用印之身分證影本，一併提供予本所，以利進行下一程序。"
'               Else
'                  g_WordAp.Selection.TypeText "四、" & custtype & "如有意續行辦理，請於隨函所附之委任書上蓋公司章及代表人章，與已用印之公司登記資料影本，一併提供予本所，以利進行下一程序。"
'               End If
'               'end 2014/10/27
'               g_WordAp.Selection.TypeParagraph
'               'modify by sonia 2014/10/27
'               'g_WordAp.Selection.TypeText "五、復審期間，初步審定暫不公告，而如屆期未提出復審，本件商標將在「" & m_GoodsName & "」商品獲准審定。"
'               g_WordAp.Selection.TypeText "五、復審期間，初步審定暫不公告，故　" & custtype & "續爭同時，可考慮以分割之方式，使無權利衝突之部分先行核准註冊（請在分割申請書上用印）。如屆期未提出復審，或復審時已申請分割，本件商標將在「" & m_GoodsName & "」商品獲准審定。"
'               'end 2014/10/27
'               g_WordAp.Selection.TypeParagraph
'               'Modify By Sindy 2009/08/11
'               'g_WordAp.Selection.TypeText "四、由於期限緊迫，請儘速與本所聯繫，以共商續行事宜。若尚有任何問題，請隨時洽詢，本所竭誠為　" & custtype & "服務。"
'               'Modify By Sindy 2012/4/23
'               'g_WordAp.Selection.TypeText "六、由於期限緊迫，請儘速於" & m_CP06 & "以前與本所聯繫，以共商續行事宜。若尚有任何問題，請隨時洽詢，本所竭誠為　" & custtype & "服務。"
'               g_WordAp.Selection.TypeText "六、由於期限緊迫，務請於" & m_DueDate & "以前與本所聯繫，以共商續行事宜。若尚有任何問題，請隨時洽詢，本所竭誠為　" & custtype & "服務。"
               g_WordAp.Selection.TypeText "三、本案續行的方式有三種："
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　1.不復審──將只在「" & m_GoodsName & "」商品／服務獲准審定。"
               g_WordAp.Selection.TypeParagraph
               '2015/8/12 modify by sonia 加入固定費用
               'g_WordAp.Selection.TypeText "　2.復審不分割，則將俟行政救濟程序確定後，大陸商標局才會針對獲准之商品／服務進行公告程序。"
               'modify by sonia 2016/9/19 28,000->29,000
               'Modified by Lydia 2017/04/27 費用改抓設定
               'g_WordAp.Selection.TypeText "　2.復審不分割，則將俟行政救濟程序確定後，大陸商標局才會針對獲准之商品／服務進行公告程序，費用為29,000元。"
               g_WordAp.Selection.TypeText "　2.復審不分割──俟行政救濟程序確定後，大陸國家知識產權局才會針對獲准之商品／服務進行公告程序，費用為" & Format("" & rsA.Fields("CFEE"), "###,##0") & "元。"
               g_WordAp.Selection.TypeParagraph
               '2015/8/12 modify by sonia 加入固定費用
               'g_WordAp.Selection.TypeText "　3.復審同時辦理分割，無權利衝突之前述部分商品／服務，將先獲准註冊。"
               'modify by sonia 2016/9/19 28,000->29,000
               'Modified by Lydia 2017/04/27 費用改抓設定
               'g_WordAp.Selection.TypeText "　3.復審同時辦理分割，無權利衝突之前述部分商品／服務，將先獲准註冊，費用為復審29,000元＋分割5,000元。"
               g_WordAp.Selection.TypeText "　3.復審同時辦理分割──無權利衝突之前述部分商品／服務，將先獲准註冊，費用為復審" & Format("" & rsA.Fields("CFEE"), "###,##0") & "元＋分割5,000元。"
               g_WordAp.Selection.TypeParagraph
               'modify by sonia 2017/8/16 取消'依法'二字
               g_WordAp.Selection.TypeText "四、本案須於" & m_CP06 & "以前向大陸國家知識產權局申請復審，且不得延期。由於本件商標與引証商標    ，故經大陸國家知識產權局認定近似，類似部分之商品不得註冊。本所以為　 。"
               g_WordAp.Selection.TypeParagraph
               If custtype = "台端" Then
                  g_WordAp.Selection.TypeText "五、續行所需文件：(1)復審：請　" & custtype & "於隨函所附之委任書上蓋章，與已用印之身分證影本；(2)分割：請　" & custtype & "在所附之分割申請書上用印。"
               Else
                  g_WordAp.Selection.TypeText "五、續行所需文件：(1)復審：請　" & custtype & "於隨函所附之評審委託書上蓋公司章及代表人章，且提供已用印之公司登記資料影本；(2)分割：請　" & custtype & "在所附之分割申請書上用印。"
               End If
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "六、因期限緊迫，務請於" & m_DueDate & "以前與本所聯繫，以共商續行事宜。"
'2014/11/27 end
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            'Add By Sindy 2009/07/02
            Case "53"
               StrSQLa = "Select CP07,CP36,NVL(NVL(CP37,CP38),CP39),NVL(NVL(CP40,CP41),CP42),CP06,CP80 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                         " AND CP10='1002' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  m_CP07 = Mid(rsA.Fields(0).Value, 1, 4) & "年" & Mid(rsA.Fields(0).Value, 5, 2) & "月" & Mid(rsA.Fields(0).Value, 7, 2) & "日"
                  m_CP06 = Mid(rsA.Fields("CP06").Value, 1, 4) & "年" & Mid(rsA.Fields("CP06").Value, 5, 2) & "月" & Mid(rsA.Fields("CP06").Value, 7, 2) & "日"
                  'Add By Sindy 2012/4/23
                  '本所期限-3工作天
                  m_DueDT = CompWorkDay(3, rsA.Fields("CP06").Value, 1)
                  m_DueDate = Mid(m_DueDT, 1, 4) & "年" & Mid(m_DueDT, 5, 2) & "月" & Mid(m_DueDT, 7, 2) & "日"
                  '2012/4/23 End
               End If
               'Modified by Lydia 2020/10/15 委由=>委託
               g_WordAp.Selection.TypeText "一、" & custtype & "委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)馬德里商標國際註冊案(本所案號" & CaseNo & ")，頃接國際局轉來" & nationname & "商標審查機關之核駁通知，隨函檢附該核駁通知書影本，請查照。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "二、駁回理由：本件商標與" & nationname & rsA.Fields(3).Value & "在類似商品上已註冊之第" & rsA.Fields(1).Value & "號「" & rsA.Fields(2).Value & "」商標近似。(註冊日：　年　月　日)指定使用於第" & IIf(rsA.Fields("CP80").Value > "" And Not IsNull(rsA.Fields("CP80").Value), rsA.Fields("CP80").Value, "　") & "類商品。"
               g_WordAp.Selection.TypeParagraph
               'Modify By Sindy 2012/4/23
               'g_WordAp.Selection.TypeText "三、本案依法須於" & m_CP07 & "向" & nationname & "特許廳提出答辯，由於本件商標與引証商標　　　。　" & custtype & "若有意爭取，請於" & m_CP06 & "以前與本所聯繫，以利提出答辯。答辯費用為新台幣　萬元整。"
               'modify by sonia 2017/8/16 取消'依法'二字
               g_WordAp.Selection.TypeText "三、本案須於" & m_CP06 & "向" & nationname & "商標審查機關提出答辯，由於本件商標與引証商標　　　。　" & custtype & "若有意爭取，請於" & m_DueDate & "以前與本所聯繫，以利提出答辯。答辯費用為新台幣　萬元整。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "四、若有任何問題，請隨時與本所聯繫，本所竭誠提供服務。"
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            '2009/07/02 End
            'Add By Sindy 2009/07/21
            Case "54"   '1201.審查報告 台->台
               'Modify By Sindy 2012/7/11
'                  StrSQLa = "Select CP07,CP36,NVL(NVL(CP37,CP38),CP39),NVL(NVL(CP40,CP41),CP42),CP06,CP80,CP08 FROM CaseProgress WHERE " & ChgCaseProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'                            " AND CP10='1201' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' "
               StrSQLa = "Select C1.CP07,C1.CP36,NVL(NVL(C1.CP37,C1.CP38),C1.CP39),NVL(NVL(C1.CP40,C1.CP41),C1.CP42),C1.CP06,C1.CP80,C1.CP08,C2.CP10,CPM03 FROM CaseProgress C1,CaseProgress C2,casepropertymap" & _
                         " WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                         " AND C1.CP10='1201' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C'" & _
                         " AND C1.CP43=C2.CP09(+)" & _
                         " AND C2.CP01=CPM01(+) AND C2.CP10=CPM02(+)"
               '2012/7/11 End
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  'modify by sonia 2019/11/20 判斷年度格式
                  m_CP07 = IIf(bolDateType, Mid(rsA.Fields(0).Value, 1, 4), Mid(rsA.Fields(0).Value, 1, 4) - 1911) & "年" & Mid(rsA.Fields(0).Value, 5, 2) & "月" & Mid(rsA.Fields(0).Value, 7, 2) & "日"
                  'Add By Sindy 2012/4/23
                  'modify by sonia 2019/11/20 判斷年度格式
                  m_CP06 = IIf(bolDateType, Mid(rsA.Fields(4).Value, 1, 4), Mid(rsA.Fields(4).Value, 1, 4) - 1911) & "年" & Mid(rsA.Fields(4).Value, 5, 2) & "月" & Mid(rsA.Fields(4).Value, 7, 2) & "日"
                  '本所期限-3工作天
                  m_DueDT = CompWorkDay(3, rsA.Fields(4).Value, 1)
                  'modify by sonia 2019/11/20 判斷年度格式
                  m_DueDate = IIf(bolDateType, Mid(m_DueDT, 1, 4), Mid(m_DueDT, 1, 4) - 1911) & "年" & Mid(m_DueDT, 5, 2) & "月" & Mid(m_DueDT, 7, 2) & "日"
                  '2012/4/23 End
               End If
               'Modify By Sindy 2012/7/11
               g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & "（第" & pa(9) & "類）商標" & Trim(rsA.Fields(8).Value) & "案，今接到智慧局之審查報告，隨函檢附智慧局" & Trim(rsA.Fields(6).Value) & "書函乙紙，請查照。"
               g_WordAp.Selection.TypeParagraph
               'Modify By Sindy 2012/4/23
               'g_WordAp.Selection.TypeText "二、本案依法須於" & m_CP07 & "以前回覆智慧局。本所建議將　　　。　" & custtype & "欲如何修正，煩請儘速與本所聯繫，以共商續行事宜。"
               g_WordAp.Selection.TypeText "二、本案依法須於" & m_CP06 & "以前回覆智慧局。需補正事項如下："
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　(1)"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　(2)"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　　" & custtype & "欲如何修正，煩請於" & m_DueDate & "以前與本所聯繫，以共商續行事宜。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "三、本案本所案號為「" & CaseNo & "」，往後查詢時，請註明本所案號以利處理。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "四、若尚有任何問題，請隨時洽詢，本所竭誠為　" & custtype & "服務。"
               '2012/7/11 End
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            Case "55"   '1201.補正通知  台->大   2010/11/5 MODIBY BY SONIA 加1702.通知修正  T-171559
               StrSQLa = "Select CP07,CP36,NVL(NVL(CP37,CP38),CP39),NVL(NVL(CP40,CP41),CP42),CP06,CP80,CP08 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                         " AND CP10 IN ('1201','1702') AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  m_CP07 = Mid(rsA.Fields(0).Value, 1, 4) & "年" & Mid(rsA.Fields(0).Value, 5, 2) & "月" & Mid(rsA.Fields(0).Value, 7, 2) & "日"
                  'Add By Sindy 2012/4/23
                  m_CP06 = Mid(rsA.Fields("CP06").Value, 1, 4) & "年" & Mid(rsA.Fields("CP06").Value, 5, 2) & "月" & Mid(rsA.Fields("CP06").Value, 7, 2) & "日"
                  '本所期限-3工作天
                  m_DueDT = CompWorkDay(3, rsA.Fields("CP06").Value, 1)
                  m_DueDate = Mid(m_DueDT, 1, 4) & "年" & Mid(m_DueDT, 5, 2) & "月" & Mid(m_DueDT, 7, 2) & "日"
                  '2012/4/23 End
               End If
               g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & AppNo & "「" & CASENAME & "」(第" & pa(9) & "類)大陸商標註冊申請案，今接到代理人轉來大陸國家知識產權局之補正通知書，謂本件商標所指定之「　　　」商品名稱不規範，隨函檢附商標註冊申請補正通知書影本乙紙，請查照。(本所案號：" & CaseNo & ")"
               g_WordAp.Selection.TypeParagraph
               'Modify By Sindy 2012/4/23
               'g_WordAp.Selection.TypeText "二、本案依法須於" & m_CP07 & "以前回覆大陸商標局。大陸代理人建議將　　　。　" & custtype & "欲如何修正，煩請儘速與本所聯繫，以共商續行事宜。"
               'modify by sonia 2017/8/16 取消'依法'二字
               g_WordAp.Selection.TypeText "二、本案須於" & m_CP06 & "以前回覆大陸國家知識產權局。大陸代理人建議將　　　。　" & custtype & "欲如何修正，煩請" & m_DueDate & "以前與本所聯繫，以共商續行事宜。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "三、若尚有任何問題，請隨時洽詢，本所竭誠為　" & custtype & "服務。"
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            '2009/07/21 End
            'ADD BY SONIA 2015/9/16  台灣案提申請意見書後之申請或分割核駁
            Case "56"   '1002
               m_CP49 = ""
               StrSQLa = "Select CP07,CP06,CP49 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                         " AND CP10='1002' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  'modify by sonia 2019/11/20 判斷年度格式
                  m_CP07 = IIf(bolDateType, Mid(rsA.Fields(0).Value, 1, 4), Mid(rsA.Fields(0).Value, 1, 4) - 1911) & "年" & Mid(rsA.Fields(0).Value, 5, 2) & "月" & Mid(rsA.Fields(0).Value, 7, 2) & "日"
                  'modify by sonia 2019/11/20 判斷年度格式
                  m_CP06 = IIf(bolDateType, Mid(rsA.Fields("CP06").Value, 1, 4), Mid(rsA.Fields("CP06").Value, 1, 4) - 1911) & "年" & Mid(rsA.Fields("CP06").Value, 5, 2) & "月" & Mid(rsA.Fields("CP06").Value, 7, 2) & "日"
                  '本所期限-3工作天
                  m_DueDT = CompWorkDay(3, rsA.Fields("CP06").Value, 1)
                  'modify by sonia 2019/11/20 判斷年度格式
                  m_DueDate = IIf(bolDateType, Mid(m_DueDT, 1, 4), Mid(m_DueDT, 1, 4) - 1911) & "年" & Mid(m_DueDT, 5, 2) & "月" & Mid(m_DueDT, 7, 2) & "日"
                  m_CP49 = "" & rsA.Fields("CP49").Value
                  GetLaw   '依條款代碼取得條款名稱caselaw
               End If
               g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之「" & CASENAME & "」(第" & pa(9) & "類)商標註冊申請乙案(本所案號" & CaseNo & ")，經提出意見書後，頃接經濟部智慧財產局" & AppNo & "審定書，謂本案與商標法規定不合，予以核駁，茲隨函檢附核駁審定正本乙份，敬請查收。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "二、本案前經智慧財產局認有違反商標法" & caselaw & "規定之嫌，雖經我方說明(提出)            但未獲採納。本案續行爭取的方向可"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　本案係有關商標法" & caselaw & "規定之爭執，我方前於意見書中主要強調　　　惟經智慧財產局認為　　　　。　" & custtype & "如有意續爭，可"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　商標法" & caselaw & "規定之適用係以　為要件之一，故我方前於意見書中，係主張　　　　　但智慧財產局仍認　　　。針對核駁理由，　" & custtype & "宜"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "三、本案依法應於經濟部智慧財產局文到次日起 30 日內提出訴願（即" & m_CP06 & "以前）否則視為結案。　" & custtype & "對上項核駁處分如有不服欲提出訴願，請於" & m_DueDate & "以前與本所聯繫，以共商提出事宜。"
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            'End 2015/9/16
         End Select
      '2008/9/24 END
      '2008/10/9 add by sonia
      Case "6"             '商爭其他C類主管機關來函
         Select Case m_Combo8
            Case "61"      '通知準備程序1203,通知言詞辯論1204(來函所選程序為勝訴)
               StrSQLa = "Select C1.CP07,C1.CP08,C1.CP64,DECODE(TM10,'000',M1.CPM03,M1.CPM04) FROM CaseProgress C1,CASEPROPERTYMAP M1,TRADEMARK WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                         " AND C1.CP10 IN ('1203','1204') AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C' AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C1.CP01=M1.CPM01(+) AND C1.CP10=M1.CPM02(+) "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  'modify by sonia 2019/11/20 判斷年度格式
                  m_CP07 = IIf(bolDateType, Mid(rsA.Fields(0).Value, 1, 4), Mid(rsA.Fields(0).Value, 1, 4) - 1911) & "年" & Mid(rsA.Fields(0).Value, 5, 2) & "月" & Mid(rsA.Fields(0).Value, 7, 2) & "日"
                  '2008/11/12 ADD BY SONIA
                  i = InStr(rsA.Fields(2).Value, ",")
                  m_TIME = Mid(rsA.Fields(2).Value, 1, i - 1)
                  m_PLACE = Mid(rsA.Fields(2).Value, i + 1)
                  '2008/11/12 END
               End If
               '2008/11/12 MODIFY BY SONIA
               'g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & casetype2 & "第" & AppNo & "號「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標" & casetype3 & "事件(本所案號" & caseno & ")，今接到智慧財產及商業法院通知，將於" & m_CP07 & rsA.Fields(2).Value & "在智慧財產及商業法院第  法庭進行" & Mid(rsA.Fields(3), 3) & "。隨函檢附智慧財產及商業法院通知書正本乙份，請查照。"
               g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & casetype2 & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標" & casetype3 & "事件(本所案號" & CaseNo & ")，今接到智慧財產及商業法院通知，將於" & m_CP07 & m_TIME & "在智慧財產及商業法院" & m_PLACE & "進行" & Mid(rsA.Fields(3), 3) & "。隨函檢附智慧財產及商業法院通知書正本乙份，請查照。"
               '2008/11/12 END
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "二、本案前經智慧財產局為有利於我方之處分，對造不服依法進行行政救濟程序，今經智慧財產及商業法院受理，並通知　" & custtype & "參加訴訟及相關程序，因事關商標權益至鉅，請速與本所聯繫，以共商續行事宜。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "三、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            Case "62"      '通知準備程序1203,通知言詞辯論1204
               StrSQLa = "Select C1.CP07,C1.CP08,C1.CP64,DECODE(TM10,'000',M1.CPM03,M1.CPM04),C2.CP10,DECODE(TM10,'000',M2.CPM03,M2.CPM04) FROM CaseProgress C1,CaseProgress C2,CASEPROPERTYMAP M1,CASEPROPERTYMAP M2,TRADEMARK WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                         " AND C1.CP10 IN ('1203','1204') AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C' AND C1.CP43=C2.CP09(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C1.CP01=M1.CPM01(+) AND C1.CP10=M1.CPM02(+) AND C2.CP01=M2.CPM01(+) AND C2.CP10=M2.CPM02(+) "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  'modify by sonia 2019/11/20 判斷年度格式
                  m_CP07 = IIf(bolDateType, Mid(rsA.Fields(0).Value, 1, 4), Mid(rsA.Fields(0).Value, 1, 4) - 1911) & "年" & Mid(rsA.Fields(0).Value, 5, 2) & "月" & Mid(rsA.Fields(0).Value, 7, 2) & "日"
                  '2008/11/12 ADD BY SONIA
                  i = InStr(rsA.Fields(2).Value, ",")
                  If i > 0 Then
                     m_TIME = Mid(rsA.Fields(2).Value, 1, i - 1)
                     m_PLACE = Mid(rsA.Fields(2).Value, i + 1)
                  End If
                  '2008/11/12 END
               End If
               Select Case rsA.Fields(4).Value
                  Case "204", "205"
                     '2008/11/12 MODIFY BY SONIA
                     'g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & casetype2 & "第" & AppNo & "號「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標" & casetype3 & "事件提起行政訴訟乙案(本所案號" & caseno & ")，今接到智慧財產及商業法院通知，將於" & m_CP07 & rsA.Fields(2).Value & "在智慧財產及商業法院第  法庭進行" & Mid(rsA.Fields(3), 3) & "。隨函檢附智慧財產及商業法院通知書正本乙份，請查照。"
                     g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & casetype2 & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標" & casetype3 & "事件提起行政訴訟乙案(本所案號" & CaseNo & ")，今接到智慧財產及商業法院通知，將於" & m_CP07 & m_TIME & "在智慧財產及商業法院" & m_PLACE & "進行" & Mid(rsA.Fields(3), 3) & "。隨函檢附智慧財產及商業法院通知書正本乙份，請查照。"
                     '2008/11/12 END
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "二、本所律師將於庭期當日代表　" & custtype & "出庭，以爭取權益。"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "三、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
                  Case "401"
                     '2008/11/12 MODIFY BY SONIA
                     'g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & casetype2 & "第" & AppNo & "號「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標" & casetype3 & "事件提起" & rsA.Fields(5) & "乙案(本所案號" & caseno & ")，今接到經濟部通知，將於" & m_CP07 & rsA.Fields(2).Value & "在經濟部進行" & Mid(rsA.Fields(3), 3) & "。隨函檢附經濟部開會通知單正本乙份，請查照。"
                     g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & casetype2 & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標" & casetype3 & "事件提起" & rsA.Fields(5) & "乙案(本所案號" & CaseNo & ")，今接到經濟部通知，將於" & m_CP07 & m_TIME & "在經濟部進行" & Mid(rsA.Fields(3), 3) & "。隨函檢附經濟部開會通知單正本乙份，請查照。"
                     '2008/11/12 END
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "二、事關商標權益至鉅，　" & custtype & "若有意親自或委託本所出席，請速聯繫，以共商續行事宜。"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "三、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
                  '2009/6/15 MODIFY BY SONIA T-065429
                  'Case "403"
                  Case Else
                  '2009/6/15 END
                     '2008/11/12 MODIFY BY SONIA
                     'g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & casetype2 & "第" & AppNo & "號「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標" & casetype3 & "事件提起" & rsA.Fields(5) & "乙案(本所案號" & caseno & ")，今接到智慧財產及商業法院通知，將於" & m_CP07 & rsA.Fields(2).Value & "在智慧財產及商業法院第  法庭進行" & Mid(rsA.Fields(3), 3) & "。隨函檢附智慧財產及商業法院通知書正本乙份，請查照。"
                     g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & casetype2 & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標" & casetype3 & "事件提起" & rsA.Fields(5) & "乙案(本所案號" & CaseNo & ")，今接到智慧財產及商業法院通知，將於" & m_CP07 & m_TIME & "在智慧財產及商業法院" & m_PLACE & "進行" & Mid(rsA.Fields(3), 3) & "。隨函檢附智慧財產及商業法院通知書正本乙份，請查照。"
                     '2008/11/12 END
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "二、本案開庭日期如上述，因事關商標權益之存廢，請速與本所聯繫，以共商續行事宜。"
                     g_WordAp.Selection.TypeParagraph
                     g_WordAp.Selection.TypeText "三、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               End Select
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            Case "63"      '通知行政上訴答辯1406
               'Modify By Sindy 2012/4/23 +C1.CP06 as CP06
               StrSQLa = "Select C1.CP07,C1.CP08,DECODE(TM10,'000',M2.CPM03,M2.CPM04),C1.CP06 as CP06 FROM CaseProgress C1,CaseProgress C2,CASEPROPERTYMAP M1,CASEPROPERTYMAP M2,TRADEMARK WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                         " AND C1.CP10='1406' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C' AND C1.CP43=C2.CP09(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C1.CP01=M1.CPM01(+) AND C1.CP10=M1.CPM02(+) AND C2.CP01=M2.CPM01(+) AND C2.CP10=M2.CPM02(+) "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  'modify by sonia 2019/11/20 判斷年度格式
                  m_CP07 = IIf(bolDateType, Mid(rsA.Fields(0).Value, 1, 4), Mid(rsA.Fields(0).Value, 1, 4) - 1911) & "年" & Mid(rsA.Fields(0).Value, 5, 2) & "月" & Mid(rsA.Fields(0).Value, 7, 2) & "日"
                  'Add By Sindy 2012/4/23
                  'modify by sonia 2019/11/20 判斷年度格式
                  m_CP06 = IIf(bolDateType, Mid(rsA.Fields("CP06").Value, 1, 4), Mid(rsA.Fields("CP06").Value, 1, 4) - 1911) & "年" & Mid(rsA.Fields("CP06").Value, 5, 2) & "月" & Mid(rsA.Fields("CP06").Value, 7, 2) & "日"
                  '2012/4/23 End
               End If
               g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & casetype2 & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標" & casetype3 & "事件(本所案號" & CaseNo & ")，前經" & rsA.Fields(2) & "並獲勝訴之判決後，對造提起上訴，隨函檢附上訴理由狀影本乙份，請查照。"
               g_WordAp.Selection.TypeParagraph
               'Modify By Sindy 2012/4/23
               'g_WordAp.Selection.TypeText "二、本案依法我方應於" & m_CP07 & "以前提出答辯狀。"
               g_WordAp.Selection.TypeText "二、本案依法我方應於" & m_CP06 & "以前提出答辯狀。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "三、有關案情本所分析以為："
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　　「對於高等行政法院之上訴，非以其違背法令為理由，不得為之。」為行政訴訟法第二四二條之規定，觀對造之上訴理由中，雖臚列若干法條，藉以證明原判決有不備理由或認定事實不依證據之失，然其主張之內容似是而非，多係原來理由之重覆陳述。我方可引據具體之相關規定，一一加以反駁，俾凸顯上訴理由之不值採納，促請最高行政法院依法駁回之，以維持既有之勝訴結果。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "四、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            Case "64"      '被異議1601,被評定1603,被廢止1605
               'modify by sonia 2019/5/27 +1619被部分廢止
               StrSQLa = "Select C1.CP08,C1.CP10,NVL(NVL(C1.CP40,C1.CP41),C1.CP42) FROM CaseProgress C1,CaseProgress C2,CASEPROPERTYMAP M1,CASEPROPERTYMAP M2,TRADEMARK WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                         " AND C1.CP10 IN ('1601','1603','1605','1619') AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C' AND C1.CP43=C2.CP09(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C1.CP01=M1.CPM01(+) AND C1.CP10=M1.CPM02(+) AND C2.CP01=M2.CPM01(+) AND C2.CP10=M2.CPM02(+) "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.Fields(1) <> "" Then
                  Select Case rsA.Fields(1)
                     Case "1601"
                        casetype2 = "註冊"
                        'casetype3 = "異議"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                        If pa(10) = "000" Then
                           casepaper = "審定書"
                        Else
                           casepaper = "裁定書"
                        End If
                     Case "1603"
                        casetype2 = "註冊"
                        If pa(10) = "000" Then
                           'casetype3 = "評定"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                           casepaper = "書"
                       Else
                           'casetype3 = "裁定"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                           '2009/11/5 MODIFY BY SONIA T-133711
                           'casepaper = "裁定書"
                           casepaper = "書"
                        End If
                     'modify by sonia 2019/5/27 +1619被部分廢止
                     Case "1605", "1619"
                        casetype2 = "註冊"
                        If pa(10) = "000" Then
                           'casetype3 = "廢止"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                           casepaper = "處分書"
                        Else
                           'casetype3 = "撤銷"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                           casepaper = "裁定書"
                        End If
                  End Select
               End If
               If rsA.Fields(1) = "1601" Then
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理註冊" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標乙案(本所案號" & CaseNo & ")，今遭" & rsA.Fields(2) & "提出" & casetype3 & "。"
               Else
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理註冊" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標乙案(本所案號" & CaseNo & ")，今遭" & rsA.Fields(2) & "申請" & casetype3 & "。"
               End If
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "二、茲寄上智慧財產局之通知書及" & casetype3 & "申請書副本各乙份，敬請查收。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "三、此案俟對造補送理由書後，本所即就案情分析。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "四、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            Case "65"      '被異議（理由）1602,被評定（理由）1604,被廢止（理由）1606
               'Modify By Sindy 2012/4/23 +C1.CP06 as CP06
               'modify by sonia 2017/9/1 +判斷Combo8 T-176979同時有1604,1606
               'modify by sonia 2019/5/27 +1620被部分廢止（理由）
               StrSQLa = "Select C1.CP07,C1.CP08,C1.CP10,NVL(NVL(C1.CP40,C1.CP41),C1.CP42),C1.CP06 as CP06 FROM CaseProgress C1,CaseProgress C2,CASEPROPERTYMAP M1,CASEPROPERTYMAP M2,TRADEMARK WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                         " AND C1.CP10 IN ('1602','1604','1606','1620') AND INSTR('" & Combo8 & "',C1.CP10)>0 AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C' AND C1.CP43=C2.CP09(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C1.CP01=M1.CPM01(+) AND C1.CP10=M1.CPM02(+) AND C2.CP01=M2.CPM01(+) AND C2.CP10=M2.CPM02(+) "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  'modify by sonia 2019/11/20 判斷年度格式
                  m_CP07 = IIf(bolDateType, Mid(rsA.Fields(0).Value, 1, 4), Mid(rsA.Fields(0).Value, 1, 4) - 1911) & "年" & Mid(rsA.Fields(0).Value, 5, 2) & "月" & Mid(rsA.Fields(0).Value, 7, 2) & "日"
                  'Add By Sindy 2012/4/23
                  'modify by sonia 2019/11/20 判斷年度格式
                  m_CP06 = IIf(bolDateType, Mid(rsA.Fields("CP06").Value, 1, 4), Mid(rsA.Fields("CP06").Value, 1, 4) - 1911) & "年" & Mid(rsA.Fields("CP06").Value, 5, 2) & "月" & Mid(rsA.Fields("CP06").Value, 7, 2) & "日"
                  '2012/4/23 End
               End If
               If rsA.Fields(2) <> "" Then
                  Select Case rsA.Fields(2)
                     Case "1602"
                        casetype2 = "註冊"
                        'casetype3 = "異議"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                        If pa(10) = "000" Then
                           casepaper = "審定書"
                        Else
                           casepaper = "裁定書"
                        End If
                     Case "1604"
                        casetype2 = "註冊"
                        If pa(10) = "000" Then
                           'casetype3 = "評定"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                           casepaper = "書"
                       Else
                           'casetype3 = "裁定"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                           '2009/11/5 MODIFY BY SONIA T-133711
                           'casepaper = "裁定書"
                           casepaper = "書"
                        End If
                     'modify by sonia 2019/5/27 +1620被部分廢止（理由）
                     Case "1606", "1620"
                        casetype2 = "註冊"
                        If pa(10) = "000" Then
                           'casetype3 = "廢止"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                           casepaper = "處分書"
                        Else
                           'casetype3 = "撤銷"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                           casepaper = "裁定書"
                        End If
                  End Select
               End If
               If rsA.Fields(2) = "1602" Then
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理註冊" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標乙案(本所案號" & CaseNo & ")，今遭" & rsA.Fields(3) & "提出" & casetype3 & "，茲即寄上" & casetype3 & "申請書副本乙份，敬請查收。"
               Else
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理註冊" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標乙案(本所案號" & CaseNo & ")，今遭" & rsA.Fields(3) & "申請" & casetype3 & "，茲即寄上" & casetype3 & "申請書副本乙份，敬請查收。"
               End If
               g_WordAp.Selection.TypeParagraph
               'Modify By Sindy 2012/4/23
               'g_WordAp.Selection.TypeText "二、本案依法我方應於" & m_CP07 & "以前提出答辯。"
               g_WordAp.Selection.TypeText "二、本案依法我方應於" & m_CP06 & "以前提出答辯。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "三、有關案情本所分析以為："
               g_WordAp.Selection.TypeParagraph
               'modify by sonia 2019/5/27 +1620被部分廢止（理由）
               If rsA.Fields(2) = "1606" Or rsA.Fields(2) = "1620" Then
                  g_WordAp.Selection.TypeText "　　　對造以系爭商標有三年以上未經使用於所指定之商品為由，認構成商標權應予廢止事由。　" & custtype & "宜檢送最近三年內將商標使用於所指定商品之證據，例如含有商標之廣告、宣傳、銷售、交易往來或進出貨等文件，俾針對對造調查報告之內容提出合理之反駁，證明無停止使用商標之事實。"
               Else
                  g_WordAp.Selection.TypeParagraph
               End If
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "四、有便煩請　" & custtype & "與本所聯絡，俾決定續行方案。　" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            Case "66"      '對方補充理由1609,發回補答辯1613   '2009/2/26加1613發回補答辯
               'Modify By Sindy 2012/4/23 +C1.CP06 as CP06
               StrSQLa = "Select C1.CP07,C1.CP08,NVL(NVL(C1.CP40,C1.CP41),C1.CP42),C2.CP10,C1.CP06 as CP06 FROM CaseProgress C1,CaseProgress C2,CASEPROPERTYMAP M1,CASEPROPERTYMAP M2,TRADEMARK WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                         " AND C1.CP10 in ('1609','1613') AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C' AND C1.CP43=C2.CP09(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C1.CP01=M1.CPM01(+) AND C1.CP10=M1.CPM02(+) AND C2.CP01=M2.CPM01(+) AND C2.CP10=M2.CPM02(+) "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  'modify by sonia 2019/11/20 判斷年度格式
                  m_CP07 = IIf(bolDateType, Mid(rsA.Fields(0).Value, 1, 4), Mid(rsA.Fields(0).Value, 1, 4) - 1911) & "年" & Mid(rsA.Fields(0).Value, 5, 2) & "月" & Mid(rsA.Fields(0).Value, 7, 2) & "日"
                  'Add By Sindy 2012/4/23
                  'modify by sonia 2019/11/20 判斷年度格式
                  m_CP06 = IIf(bolDateType, Mid(rsA.Fields("CP06").Value, 1, 4), Mid(rsA.Fields("CP06").Value, 1, 4) - 1911) & "年" & Mid(rsA.Fields("CP06").Value, 5, 2) & "月" & Mid(rsA.Fields("CP06").Value, 7, 2) & "日"
                  '2012/4/23 End
               End If
               If rsA.Fields(3) <> "" Then
                  Select Case rsA.Fields(3)
                     Case "602"
                        casetype2 = "註冊"
                        casetype3 = "異議"
                        If pa(10) = "000" Then
                           casepaper = "審定書"
                        Else
                           casepaper = "裁定書"
                        End If
                     Case "604"
                        casetype2 = "註冊"
                        If pa(10) = "000" Then
                           casetype3 = "評定"
                           casepaper = "書"
                       Else
                           casetype3 = "裁定"
                           '2009/11/5 MODIFY BY SONIA T-133711
                           'casepaper = "裁定書"
                           casepaper = "書"
                        End If
                     'modify by sonia 2019/5/27 +624部分廢止答辯
                     Case "606", "624"
                        casetype2 = "註冊"
                        If pa(10) = "000" Then
                           casetype3 = "廢止"
                           casepaper = "處分書"
                        Else
                           casetype3 = "撤銷"
                           casepaper = "裁定書"
                        End If
                  End Select
               End If
               If rsA.Fields(3) = "602" Then
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理註冊" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標(本所案號" & CaseNo & ")，前遭" & rsA.Fields(2) & "提出異議，並已答辯乙案，頃接獲商標專責機關送達之對造補充文件。隨函檢附智慧財產局之通知書及對造補充理由書副本各乙份，敬請查收。"
               Else
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理註冊" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標(本所案號" & CaseNo & ")，前遭" & rsA.Fields(2) & "申請" & casetype3 & "，並已答辯乙案，頃接獲商標專責機關送達之對造補充文件。隨函檢附智慧財產局之通知書及對造補充理由書副本各乙份，敬請查收。"
               End If
               g_WordAp.Selection.TypeParagraph
               'Modify By Sindy 2012/4/23
               'g_WordAp.Selection.TypeText "二、本案我方可於" & m_CP07 & "以前補充答辯理由及證據。"
               g_WordAp.Selection.TypeText "二、本案我方可於" & m_CP06 & "以前補充答辯理由及證據。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "三、有關案情本所分析以為："
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "四、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            Case "67"      '對方答辯1618,補證據1617
               '2008/10/20加1617補證據
               'Modify By Sindy 2012/4/23 +C1.CP06 as CP06
               StrSQLa = "Select C1.CP07,C1.CP08,C1.CP06 as CP06 FROM CaseProgress C1,CaseProgress C2,CASEPROPERTYMAP M1,CASEPROPERTYMAP M2,TRADEMARK WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                         " AND C1.CP10 IN ('1617','1618') AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C' AND C1.CP43=C2.CP09(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C1.CP01=M1.CPM01(+) AND C1.CP10=M1.CPM02(+) AND C2.CP01=M2.CPM01(+) AND C2.CP10=M2.CPM02(+) "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  'modify by sonia 2019/11/20 判斷年度格式
                  m_CP07 = IIf(bolDateType, Mid(rsA.Fields(0).Value, 1, 4), Mid(rsA.Fields(0).Value, 1, 4) - 1911) & "年" & Mid(rsA.Fields(0).Value, 5, 2) & "月" & Mid(rsA.Fields(0).Value, 7, 2) & "日"
                  'Add By Sindy 2012/4/23
                  'modify by sonia 2019/11/20 判斷年度格式
                  m_CP06 = IIf(bolDateType, Mid(rsA.Fields("CP06").Value, 1, 4), Mid(rsA.Fields("CP06").Value, 1, 4) - 1911) & "年" & Mid(rsA.Fields("CP06").Value, 5, 2) & "月" & Mid(rsA.Fields("CP06").Value, 7, 2) & "日"
                  '2012/4/23 End
               End If
               g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理註冊" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標" & casetype3 & "乙案(本所案號" & CaseNo & ")，經提出申請後，頃接獲商標專責機關送達之對造答辯書，隨函檢附智慧財產局書函及答辯理由書副本各乙份，敬請查收。"
               g_WordAp.Selection.TypeParagraph
               'Modify By Sindy 2012/4/23
               'g_WordAp.Selection.TypeText "二、本案依法我方應於" & m_CP07 & "以前補充" & casetype3 & "理由及證據。"
               g_WordAp.Selection.TypeText "二、本案依法我方應於" & m_CP06 & "以前補充" & casetype3 & "理由及證據。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "三、有關案情本所分析以為："
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "四、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            Case "68"      '智慧局答辯函1709
'2012/5/2 cancel by sonia 2012/5/1加在Command1_Click,但保留If pa(21) <> "" Then casetype2 = casetype3
'               '2009/2/26 ADD BY SONIA 卷宗性質為申請者再以案件性質判斷casetype3
'               If casetype2 = "核駁" Then
'                  StrSQLa = "Select MAX(CP05||CP10) FROM CaseProgress WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
'                            " AND CP10 IN ('101','602','1602','604','1604','606','1606') "
'                  rsA.CursorLocation = adUseClient
'                  rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'                  If rsA.Fields(0) <> "" Then
'                     Select Case Mid(rsA.Fields(0), 9)
'                        Case "602", "1602"
'                           casetype3 = "異議"
'                        Case "604", "1604"
'                           casetype3 = "評定"
'                        Case "606", "1606"
'                           casetype3 = "廢止"
'                     End Select
'                  Else
'                     If pa(21) <> "" Then casetype2 = casetype3    '2009/10/2 ADD BY SONIA T-165626
'                     casetype3 = "　　"    '中間接進來則留空白
'                  End If
'                  If rsA.State <> adStateClosed Then rsA.Close
'                  Set rsA = Nothing
'               End If
'               '2009/2/26 END
               'If pa(21) <> "" Then casetype2 = casetype3   CANCEL BY SONIA 2018/4/26 T-200842 看4/25林純貞郵件
'2012/5/2 end
               StrSQLa = "Select C1.CP08,C2.CP10 FROM CaseProgress C1,CaseProgress C2,CASEPROPERTYMAP M1,CASEPROPERTYMAP M2,TRADEMARK WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                         " AND C1.CP10='1709' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C' AND C1.CP43=C2.CP09(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C1.CP01=M1.CPM01(+) AND C1.CP10=M1.CPM02(+) AND C2.CP01=M2.CPM01(+) AND C2.CP10=M2.CPM02(+) "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.Fields(1) = "401" Then '訴願
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & casetype2 & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標" & casetype3 & "事件提起訴願乙案(本所案號" & CaseNo & ")，經提出訴願書後，業接獲經濟部智慧財產局之答辯書，茲檢附該答辯書正本乙份，敬請查收。"
               Else
                  g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & casetype2 & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標" & casetype3 & "事件提起行政訴訟乙案(本所案號" & CaseNo & ")，經提出行政訴訟起訴狀後，業接獲經濟部智慧財產局之答辯書，茲檢附該答辯書正本乙份，敬請查收。"
               End If
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "二、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
'               'Add By Sindy 2009/10/26
'               Case "69"      '變更申請案號1718
'                  StrSQLa = "Select C1.CP08,C2.CP10,TM12,C1.CP30 FROM CaseProgress C1,CaseProgress C2,CASEPROPERTYMAP M1,CASEPROPERTYMAP M2,TRADEMARK WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
'                            " AND C1.CP10='1718' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C' AND C1.CP43=C2.CP09(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C1.CP01=M1.CPM01(+) AND C1.CP10=M1.CPM02(+) AND C2.CP01=M2.CPM01(+) AND C2.CP10=M2.CPM02(+) "
'                  rsA.CursorLocation = adUseClient
'                  rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'                  g_WordAp.Selection.TypeText "一、" & custtype & "委由本所辦理之" & AppNo & "「" & CASENAME & "」" & "(第" & pa(9) & "類)商標" & casetype3 & "申請案(本所案號" & caseno & ")，頃接智慧局" & Trim(rsA.Fields(0)) & "書函，隨函檢附該函正本乙紙，請查收。"
'                  g_WordAp.Selection.TypeParagraph
'                  g_WordAp.Selection.TypeText "二、智慧局來函告以：本件商標業經　　　，應智慧局業務需要，本案將新編收文號為" & IIf(Len(Trim(rsA.Fields(2))) > 0, "0" & Trim(rsA.Fields(2)), "") & "，原申請案號為" & IIf(Len(Trim(rsA.Fields(3))) > 0, "0" & Trim(rsA.Fields(3)), "") & "，重新審辦，特此轉知。"
'                  g_WordAp.Selection.TypeParagraph
'                  g_WordAp.Selection.TypeText "三、其他若有任何質疑，亦請不吝賜教，本所當竭誠提供最佳服務。"
'                  If rsA.State <> adStateClosed Then rsA.Close
'                  Set rsA = Nothing
'               '2009/10/26 End
            'Add By Sindy 2012/4/23
            Case "6A"      '對方答辯1618  台->大
               StrSQLa = "Select C1.CP07,C1.CP08,C1.CP06 as CP06,DECODE(TM10,'000',M2.CPM03,M2.CPM04) as CPM03 FROM CaseProgress C1,CaseProgress C2,CASEPROPERTYMAP M1,CASEPROPERTYMAP M2,TRADEMARK WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                         " AND C1.CP10 IN ('1618') AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C' AND C1.CP43=C2.CP09(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C1.CP01=M1.CPM01(+) AND C1.CP10=M1.CPM02(+) AND C2.CP01=M2.CPM01(+) AND C2.CP10=M2.CPM02(+) "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  m_CP07 = Mid(rsA.Fields(0).Value, 1, 4) & "年" & Mid(rsA.Fields(0).Value, 5, 2) & "月" & Mid(rsA.Fields(0).Value, 7, 2) & "日"
                  m_CP06 = Mid(rsA.Fields("CP06").Value, 1, 4) & "年" & Mid(rsA.Fields("CP06").Value, 5, 2) & "月" & Mid(rsA.Fields("CP06").Value, 7, 2) & "日"
                  '本所期限-3工作天
                  m_DueDT = CompWorkDay(3, rsA.Fields("CP06").Value, 1)
                  m_DueDate = Mid(m_DueDT, 1, 4) & "年" & Mid(m_DueDT, 5, 2) & "月" & Mid(m_DueDT, 7, 2) & "日"
               'Modify By Sindy 2015/1/14 ex:T-177324
                  m_strCPM03 = rsA.Fields("CPM03").Value
               End If
               'g_WordAp.Selection.TypeText "一、" & custtype & "委由本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & nationname & "第" & pa(9) & "類商標" & casetype3 & rsA.Fields("CPM03").Value & "案(本所案號" & CaseNo & ")，頃接代理人轉來大陸商標局之通知，告知對造已提出答辯，我方可提供交換證據。隨函檢附該通知書及對造答辯書影本乙份，請查照。"
               'Modified by Lydia 2020/10/15 委由=>委託
               g_WordAp.Selection.TypeText "一、" & custtype & "委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & nationname & "第" & pa(9) & "類商標" & casetype3 & m_strCPM03 & "案(本所案號" & CaseNo & ")，頃接代理人轉來大陸國家知識產權局之通知，告知對造已提出答辯，我方可提供交換證據。隨函檢附該通知書及對造答辯書影本乙份，請查照。"
               '2015/1/14 END
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "二、對造於答辯書中表示：　　。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　針對對造之答辯理由，我方可　　　。此次補充理由的費用為10000元整。"
               g_WordAp.Selection.TypeParagraph
               'modify by sonia 2017/8/16 取消'依法'二字
               g_WordAp.Selection.TypeText "三、本案須於" & m_CP06 & "以前提出補充證據，" & custtype & "若有意補強理由，請儘速於" & m_DueDate & "之前與本所聯繫，以共商續行事宜。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "四、若尚有任何問題，請隨時洽詢，本所竭誠為　" & custtype & "服務。"
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            '2012/4/27 add by sonia
            Case "6B"      '對方撤回1610
               StrSQLa = "Select C1.CP08 AS CP08,C2.CP40||C2.CP41||C2.CP42 AS CP40,DECODE(SUBSTR(C2.CP09,1,1),'C',C2.CP10,C3.CP10) AS CP10,CF10 FROM CaseProgress C1,CaseProgress C2,CaseProgress C3,CASEPROPERTYMAP M1,TRADEMARK,CASEFEE WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                         " AND C1.CP10='1610' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C' AND C1.CP43=C2.CP09(+) AND C2.CP43=C3.CP09(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C1.CP01=M1.CPM01(+) AND C1.CP10=M1.CPM02(+) " & _
                         " AND C2.CP01=CF01(+) AND '" & pa(10) & "'=CF02(+) AND C2.CP10=CF03(+) "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'2012/5/2 cancel by sonia 2012/5/1加在Command1_Click
'               If rsA.Fields("CP10") <> "" Then
'                  Select Case rsA.Fields("CP10")
'                     Case "1601", "1602"
'                        casetype3 = "異議"
'                     Case "1603", "1604"
'                        If pa(10) = "000" Then
'                           casetype3 = "評定"
'                       Else
'                           casetype3 = "裁定"
'                        End If
'                     Case "1605", "1606"
'                        If pa(10) = "000" Then
'                           casetype3 = "廢止"
'                        Else
'                           casetype3 = "撤銷"
'                        End If
'                  End Select
'               End If
'2012/5/2 end
               g_WordAp.Selection.TypeText "一、" & custtype & "所有之" & casetype2 & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標，前遭" & "" & rsA.Fields("CP40") & "申請" & casetype3 & "乙案，頃接" & "" & rsA.Fields("CF10") & "來函，告知對造已撤回本件" & casetype3 & "案，隨函檢附" & "" & rsA.Fields("CF10") & "" & rsA.Fields("CP08") & "函影本乙紙，請查照。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "二、因對造撤回" & casetype3 & "案，智慧財產局將不再就本爭議案件續行審理，特此轉知。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "三、本案本所案號為「" & CaseNo & "」，往後請註明此案號，以利查詢。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "四、" & custtype & "若尚有任何問題，請隨時洽詢，本所將竭誠為　" & custtype & "提供服務。"
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            '2012/4/27 end
         End Select
      '2008/10/9 end
      Case "8"
         Select Case m_Combo8
            Case "81"      '通知復審答辯(未答辯) 台->大
               'modify by sonia 2019/5/27 +1620被部分廢止（理由）
               StrSQLa = "Select C1.CP07 as CP07,C1.CP36,NVL(NVL(C1.CP37,C1.CP38),C1.CP39),NVL(NVL(C1.CP40,C1.CP41),C1.CP42),C1.CP06 as CP06,C2.CP10 as CP10 " & _
                         "FROM CaseProgress C1,CaseProgress C2 " & _
                         "WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "' " & _
                         "AND C1.CP10='1404' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C' AND C1.CP43=C2.CP09(+) " & _
                         "AND C1.CP43>'C' AND C2.CP10 in ('1602','1604','1606','1620') "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  m_CP07 = Mid(rsA.Fields("CP07").Value, 1, 4) & "年" & Mid(rsA.Fields("CP07").Value, 5, 2) & "月" & Mid(rsA.Fields("CP07").Value, 7, 2) & "日"
                  m_CP06 = Mid(rsA.Fields("CP06").Value, 1, 4) & "年" & Mid(rsA.Fields("CP06").Value, 5, 2) & "月" & Mid(rsA.Fields("CP06").Value, 7, 2) & "日"
                  'Add By Sindy 2012/4/23
                  '本所期限-3工作天
                  m_DueDT = CompWorkDay(3, rsA.Fields("CP06").Value, 1)
                  m_DueDate = Mid(m_DueDT, 1, 4) & "年" & Mid(m_DueDT, 5, 2) & "月" & Mid(m_DueDT, 7, 2) & "日"
                  '2012/4/23 End
               End If
               If rsA.Fields("CP10") <> "" Then
                  Select Case rsA.Fields("CP10")
                     Case "1602"
                        casetype2 = "註冊"
                        'casetype3 = "異議"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                        If pa(10) = "000" Then
                           casepaper = "審定書"
                        Else
                           casepaper = "裁定書"
                        End If
                     Case "1604"
                        casetype2 = "註冊"
                        If pa(10) = "000" Then
                           'casetype3 = "評定"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                           casepaper = "書"
                       Else
                           'casetype3 = "裁定"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                           '2009/11/5 MODIFY BY SONIA T-133711
                           'casepaper = "裁定書"
                           casepaper = "書"
                        End If
                     'modify by sonia 2019/5/27 +1620被部分廢止（理由）
                     Case "1606", "1620"
                        casetype2 = "註冊"
                        If pa(10) = "000" Then
                           'casetype3 = "廢止"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                           casepaper = "處分書"
                        Else
                           'casetype3 = "撤銷"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                           casepaper = "裁定書"
                        End If
                  End Select
               End If
               g_WordAp.Selection.TypeText "一、" & custtype & "前委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & "(第" & pa(9) & "類)商標遭" & rsA.Fields(3).Value & "提出" & casetype3 & "案，今接到大陸國家知識產權局轉來" & casetype3 & "人提出復審之通知，隨函檢附該" & casetype3 & "復審通知書影本乙份，請查照。(本所案號" & CaseNo & ")"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "二、於" & casetype3 & "程序時，　" & custtype & "並未提出答辯，而大陸國家知識產權局為" & casetype3 & "不成立之決定，惟" & casetype3 & "人不服提出復審，大陸國家知識產權局乃通知　" & custtype & "就該" & casetype3 & "人之復審主張，提出答辯。"
               g_WordAp.Selection.TypeParagraph
               'Modify By Sindy 2012/4/23
               'g_WordAp.Selection.TypeText "三、本案依法於" & m_CP07 & "以前須提出答辯，" & custtype & "若有意爭取，請於" & m_CP06 & "之前儘速與本所聯繫，以共商續行事宜。"
               'modify by sonia 2017/8/16 取消'依法'二字
               g_WordAp.Selection.TypeText "三、本案須於" & m_CP06 & "以前提出答辯，　" & custtype & "若有意爭取，請於" & m_DueDate & "以前與本所聯繫，以共商續行事宜。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "四、若尚有任何問題，請隨時與洽詢，本所竭誠提供服務。"
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            Case "82"      '通知復審答辯(已答辯) 台->大
               'modify by sonia 2019/5/27 +624部分廢止答辯
               StrSQLa = "Select C1.CP07 as CP07,C1.CP36,NVL(NVL(C1.CP37,C1.CP38),C1.CP39),NVL(NVL(C1.CP40,C1.CP41),C1.CP42),C1.CP06 as CP06,C2.CP10 as CP10 " & _
                         "FROM CaseProgress C1,CaseProgress C2 " & _
                         "WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "' " & _
                         "AND C1.CP10='1404' AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C' AND C1.CP43=C2.CP09(+) " & _
                         "AND (substr(C1.CP43,1,1)='A' or substr(C1.CP43,1,1)='B') AND C2.CP10 in ('602','604','606','624') "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  m_CP07 = Mid(rsA.Fields("CP07").Value, 1, 4) & "年" & Mid(rsA.Fields("CP07").Value, 5, 2) & "月" & Mid(rsA.Fields("CP07").Value, 7, 2) & "日"
                  m_CP06 = Mid(rsA.Fields("CP06").Value, 1, 4) & "年" & Mid(rsA.Fields("CP06").Value, 5, 2) & "月" & Mid(rsA.Fields("CP06").Value, 7, 2) & "日"
                  'Add By Sindy 2012/4/23
                  '本所期限-3工作天
                  m_DueDT = CompWorkDay(3, rsA.Fields("CP06").Value, 1)
                  m_DueDate = Mid(m_DueDT, 1, 4) & "年" & Mid(m_DueDT, 5, 2) & "月" & Mid(m_DueDT, 7, 2) & "日"
                  '2012/4/23 End
               End If
               If rsA.Fields("CP10") <> "" Then
                  Select Case rsA.Fields("CP10")
                     Case "602"
                        casetype2 = "註冊"
                        'casetype3 = "異議"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                        If pa(10) = "000" Then
                           casepaper = "審定書"
                        Else
                           casepaper = "裁定書"
                        End If
                     Case "604"
                        casetype2 = "註冊"
                        If pa(10) = "000" Then
                           'casetype3 = "評定"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                           casepaper = "書"
                       Else
                           'casetype3 = "裁定"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                           '2009/11/5 MODIFY BY SONIA T-133711
                           'casepaper = "裁定書"
                           casepaper = "書"
                        End If
                     'modify by sonia 2019/5/27 +624部分廢止答辯
                     Case "606", "624"
                        casetype2 = "註冊"
                        If pa(10) = "000" Then
                           'casetype3 = "廢止"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                           casepaper = "處分書"
                        Else
                           'casetype3 = "撤銷"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                           casepaper = "裁定書"
                        End If
                  End Select
               End If
               'Modified by Lydia 2020/10/15 委由=>委託
               g_WordAp.Selection.TypeText "一、" & custtype & "委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & nationname & "(第" & pa(9) & "類)商標遭" & rsA.Fields(3).Value & "提出" & casetype3 & "案(本所案號" & CaseNo & ")，今接到大陸國家知識產權局轉來" & casetype3 & "人提出復審之通知，隨函檢附該" & casetype3 & "復審通知書影本乙份，請查照。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "二、" & casetype3 & "人於復審程序再度強調　　　，　" & custtype & "應答辯力爭之。"
               g_WordAp.Selection.TypeParagraph
               'Modify By Sindy 2012/4/23
               'g_WordAp.Selection.TypeText "三、本案依法於" & m_CP07 & "以前須提出答辯，　" & custtype & "若有意爭取，請儘速於" & m_CP06 & "之前與本所聯繫，以共商續行事宜。"
               'modify by sonia 2017/8/16 取消'依法'二字
               g_WordAp.Selection.TypeText "三、本案須於" & m_CP06 & "以前提出答辯，　" & custtype & "若有意爭取，請於" & m_DueDate & "以前與本所聯繫，以共商續行事宜。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "四、若尚有任何問題，請隨時與洽詢，本所竭誠提供服務。"
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            Case "83"      '被異議,被評定 台->大
               'modify by sonia 2017/9/1 +判斷Combo8 因台灣案T-176979同時有1604,1606
               'modify by sonia 2019/5/27 +1620被部分廢止（理由）
               StrSQLa = "Select C1.CP07 as CP07,C1.CP08,C1.CP10,NVL(NVL(C1.CP40,C1.CP41),C1.CP42),NVL(NVL(C1.CP37,C1.CP38),C1.CP39),C1.CP36,C1.CP06 as CP06 " & _
                         "FROM CaseProgress C1,CaseProgress C2,CASEPROPERTYMAP M1,CASEPROPERTYMAP M2,TRADEMARK " & _
                         "WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "' " & _
                         "AND C1.CP10 IN ('1602','1604','1606','1620') AND INSTR('" & Combo8 & "',C1.CP10)>0 AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C' AND C1.CP43=C2.CP09(+) " & _
                         "AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) " & _
                         "AND C1.CP01=M1.CPM01(+) AND C1.CP10=M1.CPM02(+) AND C2.CP01=M2.CPM01(+) AND C2.CP10=M2.CPM02(+) "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  m_CP07 = Mid(rsA.Fields("CP07").Value, 1, 4) & "年" & Mid(rsA.Fields("CP07").Value, 5, 2) & "月" & Mid(rsA.Fields("CP07").Value, 7, 2) & "日"
                  m_CP06 = Mid(rsA.Fields("CP06").Value, 1, 4) & "年" & Mid(rsA.Fields("CP06").Value, 5, 2) & "月" & Mid(rsA.Fields("CP06").Value, 7, 2) & "日"
                  'Add By Sindy 2012/4/23
                  '本所期限-3工作天
                  m_DueDT = CompWorkDay(3, rsA.Fields("CP06").Value, 1)
                  m_DueDate = Mid(m_DueDT, 1, 4) & "年" & Mid(m_DueDT, 5, 2) & "月" & Mid(m_DueDT, 7, 2) & "日"
                  '2012/4/23 End
               End If
               If rsA.Fields(2) <> "" Then
                  Select Case rsA.Fields(2)
                     Case "1602"
                        casepaper = ""
                        casecopy = "異議理由書"
                     Case "1604"
                        casetype3 = "無效宣告請求"
                        casepaper = "書"
                        casecopy = "請求書"
                  End Select
               End If
               'Modified by Lydia 2020/10/15 委由=>委託
               g_WordAp.Selection.TypeText "一、" & custtype & "委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & nationname & "(第" & pa(9) & "類)商標註冊申請案(本所案號" & CaseNo & ")，頃接今接到代理人轉來大陸國家知識產權局之商標" & casetype3 & casepaper & "的答辯通知，謂本件商標遭" & rsA.Fields(3).Value & "提出" & casetype3 & "，隨函檢附該答辯通知書及" & casecopy & "影本乙份，請查照。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "二、" & casetype3 & "理由：" & rsA.Fields(3).Value & "以本件商標與其註冊在先之第" & rsA.Fields(5).Value & "號「" & rsA.Fields(4).Value & "」商標相同及與　　　商標近似，且指定用相類似之服務，有違商標法的規定。"
               g_WordAp.Selection.TypeParagraph
               'Modify By Sindy 2012/4/23
               'g_WordAp.Selection.TypeText "三、本案依法須於" & m_CP07 & "以前提出答辯，由於　　　，本案應可以此理由爭取之。"
               'modify by sonia 2017/8/16 取消'依法'二字
               g_WordAp.Selection.TypeText "三、本案須於" & m_CP06 & "以前提出答辯，由於　　　，本案應可以此理由爭取之。"
               g_WordAp.Selection.TypeParagraph
               'Modify By Sindy 2012/4/23
               'g_WordAp.Selection.TypeText "四、" & custtype & "若有意提出答辯，請於" & m_CP06 & "之前儘速與本所聯繫，以共商續行事宜。"
               g_WordAp.Selection.TypeText "四、" & custtype & "若有意提出答辯，請於" & m_DueDate & "以前與本所聯繫，以共商續行事宜。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "五、若尚有任何問題，請隨時與洽詢，本所竭誠為　" & custtype & "服務。"
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            Case "84"      '被撤銷(三年未使用) 台->大
               'modify by sonia 2019/5/27 +1620被部分廢止（理由）
               StrSQLa = "Select C1.CP07 as CP07,C1.CP08,C1.CP10,NVL(NVL(C1.CP40,C1.CP41),C1.CP42),C1.CP06 as CP06 " & _
                         "FROM CaseProgress C1,CaseProgress C2,CASEPROPERTYMAP M1,CASEPROPERTYMAP M2,TRADEMARK " & _
                         "WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "' " & _
                         "AND C1.CP10 IN ('1606','1620') AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C' AND C1.CP43=C2.CP09(+) " & _
                         "AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) " & _
                         "AND C1.CP01=M1.CPM01(+) AND C1.CP10=M1.CPM02(+) AND C2.CP01=M2.CPM01(+) AND C2.CP10=M2.CPM02(+) "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  m_CP07 = Mid(rsA.Fields("CP07").Value, 1, 4) & "年" & Mid(rsA.Fields("CP07").Value, 5, 2) & "月" & Mid(rsA.Fields("CP07").Value, 7, 2) & "日"
                  m_CP06 = Mid(rsA.Fields("CP06").Value, 1, 4) & "年" & Mid(rsA.Fields("CP06").Value, 5, 2) & "月" & Mid(rsA.Fields("CP06").Value, 7, 2) & "日"
                  'Add By Sindy 2012/4/23
                  '本所期限-3工作天
                  m_DueDT = CompWorkDay(3, rsA.Fields("CP06").Value, 1)
                  m_DueDate = Mid(m_DueDT, 1, 4) & "年" & Mid(m_DueDT, 5, 2) & "月" & Mid(m_DueDT, 7, 2) & "日"
                  '2012/4/23 End
               End If
               'Modified by Lydia 2020/10/15 委由=>委託
               g_WordAp.Selection.TypeText "一、" & custtype & "委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & nationname & "(第" & pa(9) & "類)商標註冊申請案(本所案號" & CaseNo & ")，頃接代理人轉來大陸國家知識產權局之通知，謂本件商標遭" & rsA.Fields(3).Value & "提出商標三年未使用撤銷，隨函檢附該通知書影本乙份，請查照。"
               g_WordAp.Selection.TypeParagraph
               'modify by sonia 2022/6/14 取消開頭之 本件商標遭申請撤銷，
               g_WordAp.Selection.TypeText "二、大陸國家知識產權局通知必須檢附　　　間本件商標於大陸的使用証明，如：廣告、參展証明、販售商品發票、經銷商契約‧‧‧等正本文件，如未提出於大陸的使用証明，則本件商標將遭撤銷註冊。"
               g_WordAp.Selection.TypeParagraph
               'Modify By Sindy 2012/4/23
               'g_WordAp.Selection.TypeText "三、本案依法須於" & m_CP07 & "以前提出答辯，　" & custtype & "若有意爭取，請儘速於" & m_CP06 & "之前與本所聯繫，以共商續行事宜。"
               'modify by sonia 2017/8/16 取消'依法'二字
               g_WordAp.Selection.TypeText "三、本案須於" & m_CP06 & "以前提出答辯，　" & custtype & "若有意爭取，請於" & m_DueDate & "以前與本所聯繫，以共商續行事宜。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "四、若尚有任何問題，請隨時與洽詢，本所竭誠為　" & custtype & "服務。"
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
            'Add By Sindy 2009/08/21
            Case "85"      '被異議（理由）1602,被評定（理由）1604,被廢止（理由）1606 --- TF(馬德里)
               'modify by sonia 2017/9/1 +判斷Combo8 因台灣案T-176979同時有1604,1606
               'modify by sonia 2019/5/27 +1620被部分廢止（理由）
               StrSQLa = "Select C1.CP07,C1.CP08,C1.CP10,NVL(NVL(C1.CP40,C1.CP41),C1.CP42),C1.CP06,C1.CP40,C1.CP41 FROM CaseProgress C1,CaseProgress C2,CASEPROPERTYMAP M1,CASEPROPERTYMAP M2,TRADEMARK WHERE C1.CP01='" & pa(1) & "' AND C1.CP02='" & pa(2) & "' AND C1.CP03='" & pa(3) & "' AND C1.CP04='" & pa(4) & "'" & _
                         " AND C1.CP10 IN ('1602','1604','1606','1620') AND INSTR('" & Combo8 & "',C1.CP10)>0 AND C1.CP27 IS NULL AND C1.CP57 IS NULL AND C1.CP09>'C' AND C1.CP43=C2.CP09(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND C1.CP01=M1.CPM01(+) AND C1.CP10=M1.CPM02(+) AND C2.CP01=M2.CPM01(+) AND C2.CP10=M2.CPM02(+) "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  m_CP07 = Mid(rsA.Fields(0).Value, 1, 4) & "年" & Mid(rsA.Fields(0).Value, 5, 2) & "月" & Mid(rsA.Fields(0).Value, 7, 2) & "日"
                  m_CP06 = Mid(rsA.Fields(4).Value, 1, 4) & "年" & Mid(rsA.Fields(4).Value, 5, 2) & "月" & Mid(rsA.Fields(4).Value, 7, 2) & "日"
                  'Add By Sindy 2012/4/23
                  '本所期限-3工作天
                  m_DueDT = CompWorkDay(3, rsA.Fields(4).Value, 1)
                  m_DueDate = Mid(m_DueDT, 1, 4) & "年" & Mid(m_DueDT, 5, 2) & "月" & Mid(m_DueDT, 7, 2) & "日"
                  '2012/4/23 End
               End If
               If rsA.Fields(2) <> "" Then
                  Select Case rsA.Fields(2)
                     Case "1602"
                        casetype2 = "註冊"
                        'casetype3 = "異議"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                        If pa(10) = "000" Then
                           casepaper = "審定書"
                        Else
                           casepaper = "裁定書"
                        End If
                     Case "1604"
                        casetype2 = "註冊"
                        If pa(10) = "000" Then
                           'casetype3 = "評定"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                           casepaper = "書"
                       Else
                           'casetype3 = "裁定"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                           '2009/11/5 MODIFY BY SONIA T-133711
                           'casepaper = "裁定書"
                           casepaper = "書"
                        End If
                     'modify by sonia 2019/5/27 +1620被部分廢止（理由）
                     Case "1606", "1620"
                        casetype2 = "註冊"
                        If pa(10) = "000" Then
                           'casetype3 = "廢止"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                           casepaper = "處分書"
                        Else
                           'casetype3 = "撤銷"      '2012/5/1 cancel by sonia 2012/5/1加在Command1_Click
                           casepaper = "裁定書"
                        End If
                  End Select
               End If
               'Modified by Lydia 2020/10/15 委由=>委託
               g_WordAp.Selection.TypeText "　　　　" & custtype & "委託本所辦理之" & AppNo & "「" & CASENAME & "」" & casetype4 & nationname & "(第" & pa(9) & "類)商標註冊申請案(本所案號" & CaseNo & ")，頃接代理人來函通知，德商" & rsA.Fields(6) & "(以下稱" & rsA.Fields(5) & ")已對本商標提出" & casetype3 & "。隨函檢附" & nationname & "大陸國家知識產權局核發之正式" & casetype3 & "通知函、對造提出之" & casetype3 & "理由書及下列據以" & casetype3 & "商標之資料供　" & custtype & "參考："
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　商標名稱："
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　註冊號數："
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　申請日期：　年　月　日"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　註冊日：　年　月　日"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　" & nationname & "最早使用日期：　年　月　日"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　指定商品：第　類之"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　目前進度：註冊有效"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　　　因" & nationname & "商標審理暨訴願委員已受理" & casetype3 & "，故已設定" & casetype3 & "案預定進行之程序及時間表如下："
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeParagraph
               'modify by sonia 2017/8/16 所有台->外撰寫信函都不放法定期限,故改m_CP06
               'g_WordAp.Selection.TypeText "　　　　申請人(即　" & custtype & ")提" & casetype3 & "答辯之法定期限為" & m_CP07
               g_WordAp.Selection.TypeText "　　　　申請人(即　" & custtype & ")提" & casetype3 & "答辯之期限為" & m_CP06
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　　　證據調查會議截止日　　　　　年　月　日"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　　　證據調查起始日　　　　　　　年　月　日"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　　　首次揭露截止日　　　　　　　年　月　日"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　　　專家揭露截止日　　　　　　　年　月　日"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　　　證據調查截止日　　　　　　　年　月　日"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　　　原告初審揭露　　　　　　　　年　月　日"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　　　原告30天審問期間截止日　　　年　月　日"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　　　被告初審揭露　　　　　　　　年　月　日"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　　　被告30天審問期間截止日　　　年　月　日"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　　　原告反駁對方的揭露起始日　　年　月　日"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　　　原告15天反駁期間截止日　　　年　月　日"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　　　以上法定期限除" & casetype3 & "答辯期限不能延期外，其餘皆可延期，惟須雙方同時提出申請。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　　　關於本案之續行，　" & custtype & "可選擇下列二個方案之一處理："
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeParagraph
               'Modify By Sindy 2012/4/23
               'g_WordAp.Selection.TypeText "一、在法定期限" & m_CP07 & "前提" & casetype3 & "答辯爭取，逾期未提答辯將導致案件失效。在我方提答辯後，若" & casetype3 & "人堅持續行" & casetype3 & "程序，則" & casetype3 & "案將正式進入訴訟階段，因" & nationname & casetype3 & "程序之進行係比照民事訴訟程序，由雙方當事人以書面或口頭辯論方式逐步進行攻防，程序進行中若有一方放棄答辯，即視同敗訴，而不論商標是否確實構成近似，整個" & casetype3 & "程序所需時間可能長達數年，費用亦可能高達新台幣數佰萬元，故請慎重考慮是否續行" & casetype3 & "答辯程序。提初步" & casetype3 & "答辯之費用為新台幣　　　元整；若提答辯後，對造仍堅持續行" & casetype3 & "程序，則必須進行證據調查，證據調查程序之費用另計（代理人預估僅證據調查程序之費用至少需新台幣伍拾萬元）。"
               g_WordAp.Selection.TypeText "一、在期限" & m_CP06 & "前提" & casetype3 & "答辯爭取，逾期未提答辯將導致案件失效。在我方提答辯後，若" & casetype3 & "人堅持續行" & casetype3 & "程序，則" & casetype3 & "案將正式進入訴訟階段，因" & nationname & casetype3 & "程序之進行係比照民事訴訟程序，由雙方當事人以書面或口頭辯論方式逐步進行攻防，程序進行中若有一方放棄答辯，即視同敗訴，而不論商標是否確實構成近似，整個" & casetype3 & "程序所需時間可能長達數年，費用亦可能高達新台幣數佰萬元，故請慎重考慮是否續行" & casetype3 & "答辯程序。提初步" & casetype3 & "答辯之費用為新台幣　　　元整；若提答辯後，對造仍堅持續行" & casetype3 & "程序，則必須進行證據調查，證據調查程序之費用另計（代理人預估僅證據調查程序之費用至少需新台幣伍拾萬元）。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "二、嘗試與對造進行協商撤回" & casetype3 & "，惟因本案提" & casetype3 & "答辯之法定期限不能延期，故若　" & custtype & "希望經由協商解決商標爭議，仍須先提答辯維持申請效力以爭取協商時間。協商可由　" & custtype & "自行處理亦可委託代理人處理。另　" & custtype & "亦可提供本商標之實際使用態樣及商品照片、型錄等供對造參考。若　" & custtype & "希望透過代理人與對造進行協商，則請代理人發函詢問對造協商意願之費用為新台幣　　　　元整，包括代理人撰寫來信通知被" & casetype3 & "時已衍生之服務費，但不包括後續之協商及延期費用。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　　　因據以" & casetype3 & "商標主要用在　　　　，故若　" & custtype & "同意進行協商，則代理人建議限定本商標之指定商品為僅包括下列產品，以期與對造之產品有所區隔，則雙方達成協議之機會較大："
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "三、不提答辯，讓本案因逾期未提答辯而喪失申請效力，惟因代理人撰寫來信通知他人延期提" & casetype3 & "及通知被" & casetype3 & "皆已衍生服務費，故若　" & custtype & "選擇此項方案續行，本所須酌收代理人服務費新台幣　　　元整，以支付" & nationname & "代理人請款。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeParagraph
               'Modify By Sindy 2012/4/23
               'g_WordAp.Selection.TypeText "　　　　" & custtype & "欲如何續行本案，敬請務必於" & m_CP06 & "前惠示本所，以利適時通知代理人處理。其他若有任何需要本所服務的地方，請不吝賜教，本所當竭誠為　" & custtype & "提供服務。"
               g_WordAp.Selection.TypeText "　　　　" & custtype & "欲如何續行本案，敬請務必於" & m_DueDate & "以前惠示本所，以利適時通知代理人處理。其他若有任何需要本所服務的地方，請不吝賜教，本所當竭誠為　" & custtype & "提供服務。"
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeParagraph
               g_WordAp.Selection.TypeText "　　其他若有任何質疑，請不吝賜教，本所當竭誠提供最佳服務。"
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
         End Select
   End Select
End Sub

'Modify by Morgan 2009/8/13 PCT檢索報告報告定稿PCT成員國改3個
'Modify by Morgan 2008/7/17 改開窗信封用的信紙格式
'Modify by Morgan 2016/11/10 +pIsCopy:是否為副本
'Modify by Morgan 2017/3/14 +pNoCopy:是否列印副本收件人
Private Sub WordChinese(Optional pIsCopy As Boolean = False, Optional pCuNo As String, Optional pContNo As String, Optional pNoCopy As Boolean)
Dim stReceiver As String '收件人
Dim stAddr As String '地址
Dim stZip As String '郵遞區號
Dim stContact As String '接洽人
Dim iLineCount As Integer '行數
Dim stApp1Title As String '稱謂
Dim iPicNo As Integer, stFileName As String 'Add by Morgan 2010/12/13
Dim iPicNo2 As Integer 'Add by Morgan 2011/5/10 信尾圖檔代碼
Dim oShape
Dim stMultiApp As String '多人申請 Added by Morgan 2013/7/5
'Add By Sindy 2013/11/14
Dim strText As String, intLenText As Integer
'2013/11/14 END
Dim strLetterNo As String '發文號 Added by Morgan 2014/12/8
'Added by Morgan 2016/11/11
Dim stCC As String '副本收受人
Dim stAppName As String '申請人(收件人不一定是申請人 Ex.副本)
'end 2016/11/11
Dim lngPaper As Long 'Added by Morgan 2018/10/29 ｅ化客戶是否寄紙本
Dim stCP09 As String 'Added by Morgan 2019/2/14 信函收文號
Dim bolECust As Boolean 'Added by Morgan 2021/12/2 是否全E化客戶
Dim bolSalesHandle As Boolean 'Added by Morgan 2024/11/11 是否業務自行處理
Dim strData As String 'Added by Morgan 2025/4/9
   
'*******2019/11/19 統一台灣案用民國年月日,但外至台改用西元年月日,非台灣案用西元年月日
   'Added by Morgan 2013/7/5
   m_MySt(1) = pa(1)
   m_MySt(2) = pa(2)
   m_MySt(3) = pa(3)
   m_MySt(4) = pa(4)
   m_SysKind = CheckSys(pa(1))
   SetLetterSt
   stMultiApp = ExceptFieldData2("多人申請")
   'end 2013/7/5
      
      iLineCount = 0
      
      If Option2.Value Then
         stReceiver = Trim(fa(0))
         stAddr = Trim(fa(16))
         
      ElseIf Option6.Value Then
         stReceiver = Trim(cfa(0))
         stAddr = Trim(cfa(16))
                  
      ElseIf Option3.Value Then
'Modify by Morgan 2008/8/8 改呼叫共用函數
'         If m_CU104 <> "" Then
'　　　　　  stReceiver = m_CU104
'         Else
'　　　　　  stReceiver = Trim(cu(0))
'         End If
'         'end 2008/8/6
'         If m_CU80 = "" Then
'　　　　　  stAddr = Trim(cu(16))
'　　　　　  stZip = Trim(m_Zip)
'　　　　　  stContact = Trim(m_Contact)
'         Else
'　　　　　  stAddr = String(20, " ")
'　　　　　  stZip = String(17, " ")
'　　　　　  stContact = String(3, " ")
'         End If
         'Modified by Morgan 2020/10/16 +傳入stAppName(因若有更名此參數會回傳舊名稱,收件人則固定抓最新名稱)
         Call PUB_GetAddrRef("", Text1, Text2, Text3, Text4, stReceiver, stContact, stZip, stAddr, , , , stAppName, "1")
         stApp1Title = PUB_GetAppTitle(m_CustNo(1))
         If stApp1Title <> "" Then
            stApp1Title = "　" & stApp1Title
         End If
         
         If lngPaper = vbNo Then stZip = "": stAddr = "" 'Added by Morgan 2018/10/29 不寄紙本時不印郵遞區號及地址
         
'end 2008/8/8
      End If
      
      'Added by Morgan 2016/11/11
      If stAppName = "" Then stAppName = stReceiver
      If Option3.Value Then
         stCC = ExceptFieldData2("副本收受人")
         If pIsCopy Then
            Call PUB_GetAddrRef(pCuNo, , , , , stReceiver, stContact, stZip, stAddr, , pContNo, , , "1")
            If lngPaper = vbNo Then stZip = "": stAddr = "" 'Added by Morgan 2018/11/21 不寄紙本時不印郵遞區號及地址
         End If
      End If
      'end 2016/11/11
      
      strData = "致：" & stAppName & stApp1Title 'Added by Morgan 2025/4/9
      

'Added by Morgan 2020/3/25
If strSrvDate(1) >= 智慧所更名日 Then
   m_CompNo = PUB_GetReceiptComp(pa(1), pa(2), pa(3), pa(4), True)
   PUB_GetLetterPicID m_CompNo, pa(1), iPicNo, iPicNo2, 1, False, m_Dept
Else
'end 2020/3/25

   'Add by Morgan 2010/12/13 加可印信頭
   'Modifided by Morgan 2013/12/30 改都要判斷,且改回傳公司別
   m_CompNo = PUB_GetReceiptComp(pa(1), pa(2), pa(3), pa(4))
   '有設定
   '專利商標
   If m_CompNo = "T" Then
      iPicNo = 12
      iPicNo2 = 11
   '智權公司
   ElseIf m_CompNo = "J" Then
      iPicNo = 28
      'Modified by Morgan 2015/11/18
      '統一用10樓的傳真--陳鳳英
      'If Text1.Text = "CFT" Then
      '   iPicNo2 = 29 'fax:25090804
      'Else
         iPicNo2 = 31 'fax:25011666
      'End If
      'end 2015/11/18

   '未設定
   Else
      '專利商標
      'modify by sonia 2014/4/28
      'If Text1.Text = "T" Then
      'Modified by Morgan 2015/2/10 +S
      If InStr(Text1.Text, "T") > 0 Or Text1.Text = "S" Then
         iPicNo = 12
      '專利法律
      Else
         iPicNo = 10
      End If
      iPicNo2 = 11
   End If
   'end 2013/12/30
   
End If 'Added by Morgan 2020/3/25
   
   bolRetry = False
   
On Error GoTo ERRORSECTION1
   
    If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
    g_WordAp.Documents.add
    With g_WordAp
      'Add by Morgan 2013/1/8
      '切換為整頁模式,信頭才會正常顯示
      If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
         .ActiveWindow.ActivePane.View.Type = wdPageView
      Else
         .ActiveWindow.View.Type = wdPageView
      End If
      'end 2013/1/8
      
      .Selection.Font.Name = "標楷體"
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.Orientation = wdTextOrientationHorizontal
      .Selection.Font.Size = 14
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(2)
      'Modify by Morgan 2008/7/3
      '.Selection.PageSetup.TopMargin = .CentimetersToPoints(3.53)
      '.Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
      
'Modify by Morgan 2011/5/10 改用黑白的信頭(尾)格式
'      .Selection.PageSetup.TopMargin = .CentimetersToPoints(4.1)
'      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(2.5)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(4.2)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(3)
'end 2011/5/10

      'end 2008/7/3
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
      .Selection.ParagraphFormat.DisableLineHeightGrid = True
      'Add by Morgan 2008/7/17 配合新的開窗定稿改固定行高
      .Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
      .Selection.ParagraphFormat.LineSpacing = 15
      'end 2008/7/17
      
      'Add by Morgan 2010/12/13
      '信函信頭
      If txtLetterHead <> "N" Then
         If PUB_ReadDB2File(stFileName, iPicNo) = True Then
            '插入圖片檔案
            'Modify by Morgan 2011/5/10 改放在頁首(尾)才不會影響輸入操作且跳頁也不必重複插入信頭(尾)
            .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
            PUB_AddConfidential .Selection, strData 'Added by Morgan 2025/4/9
            
            Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
            oShape.ZOrder 4
            oShape.LockAnchor = True
            oShape.LockAspectRatio = -1
            oShape.WrapFormat.Type = 5 'Added by Morgan 2025/4/9
            oShape.Width = .CentimetersToPoints(21)
            oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
            oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
            oShape.Left = .CentimetersToPoints(0)
            'Added by Morgan 2020/3/26 統一
            If strSrvDate(1) >= 智慧所更名日 Then
               oShape.Top = .CentimetersToPoints(0)
            Else
               oShape.Top = .CentimetersToPoints(0.5)
            End If
            If PUB_ReadDB2File(stFileName, iPicNo2) = True Then
               .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
               Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
               oShape.ZOrder 4
               oShape.LockAnchor = True
               oShape.LockAspectRatio = -1
               oShape.Width = .CentimetersToPoints(21)
               oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
               oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
               oShape.Left = .CentimetersToPoints(0)
               'Added by Morgan 2020/3/26
               If strSrvDate(1) >= 智慧所更名日 Then
                  oShape.Top = .CentimetersToPoints(27.2)
               Else
                  oShape.Top = .CentimetersToPoints(27)
               End If
               
            End If
            .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
            'End 2011/5/10
            .Selection.EndKey Unit:=wdStory
         End If
      End If
      
      'Modified by Morgan 2018/10/29 全E化客戶要印指定信箱
      '.Selection.TypeParagraph
      If strSrvDate(1) >= e化客戶啟用日 And Option3.Value Then
         'Added by Morgan 2018/11/21
         '副本E化要判斷收受人
         'Modified by Morgan 2024/11/11 +bolSalesHandle
         If pIsCopy Then
            If PUB_ChkECust(pCuNo, Text1, strExc(0), intI, bolSalesHandle) = True Then
               'bolECust = True 'Added by Morgan 2021/12/2 Removed by Morgan 2022/3/9 移到下面
               .Selection.TypeText "E-mail: " & strExc(0)
               If intI = 1 Then
                  bolECust = True 'Added by Morgan 2022/3/9
                  lngPaper = vbNo
                  .Selection.TypeParagraph 'Added by Morgan 2024/1/31 少一行表圖會太貼內文 Ex:CFP-032391
               Else
                  .Selection.TypeParagraph
               End If
            Else
               .Selection.TypeParagraph
            End If
         Else
         'end 2018/11/21
            If PUB_ChkECust(m_CustNo(1), Text1, strExc(0), intI, bolSalesHandle) = True Then
               'bolECust = True 'Added by Morgan 2021/12/2 Removed by Morgan 2022/3/9 移到下面
               .Selection.TypeText "E-mail: " & strExc(0)
               If intI = 1 Then
                  bolECust = True 'Added by Morgan 2021/12/2
                  'Modified by Morgan 2021/11/18 承辦人寫信都不會有實體，改一律不印地址
                  'lngPaper = MsgBox(m_CustNo(1) & " " & Combo3 & " 為全E化客戶！" & vbCrLf & "本信函是否需寄送【紙本】？" & vbCrLf & vbCrLf & "※有實體文件(例如收據、證書…)請選【是】" & vbCrLf & "※選【否】將不印【郵遞區號及地址】", vbYesNo + vbDefaultButton2 + vbQuestion, "ｅ化客戶提醒")
                  MsgBox m_CustNo(1) & " " & Combo3 & " 為全E化客戶！將不印【郵遞區號及地址】及【掛號】並於下款加印【公司章】。", vbInformation, "全E化客戶提醒"
                  lngPaper = vbNo
                  'end 2021/11/18
                  .Selection.TypeParagraph 'Added by Morgan 2024/1/31 少一行表圖會太貼內文 Ex:CFP-032391
               Else
                  'Added by Morgan 2024/11/11 時間點有點早且專利商標流程不一，先保留不上
                  'If bolSalesHandle Then
                  '   MsgBox m_CustNo(1) & " " & Combo3 & " 客戶狀態為【業務自行處理】且為【半E化】！" & vbCrLf & vbCrLf & "客戶函完成(判發)後不需列印紙本！", vbInformation, "半E化客戶提醒"
                  'End If
                  'end 2024/11/11
                  .Selection.TypeParagraph
               End If
            Else
               .Selection.TypeParagraph
            End If
         End If
      Else
         .Selection.TypeParagraph
      End If
      'end 2018/10/29
      
      .Selection.TypeParagraph 'Add by Morgan 2008/6/11 CFT 信頭比較高
      'Add By Sindy 2013/11/14 +因加右代表圖,所以調整其他欄位的位置
      'Add By Sindy 2014/3/5 P,CFP也要加右代表圖
      'Modified by Lydia 2018/08/22 所有 T字頭的系統類別, 都加代表圖
      'If Text1 <> "T" And Text1 <> "CFT" And Text1 <> "P" And Text1 <> "CFP" Then
      If Left(Text1, 1) <> "T" And Text1 <> "CFT" And Text1 <> "P" And Text1 <> "CFP" Then
         .Selection.TypeParagraph
      End If
      'Remove by Morgan 2008/7/3 上邊界增加
      '.Selection.TypeParagraph
      '.Selection.TypeParagraph
      
      'Added by Morgan 2025/3/19 除錯用
      If g_LetterDebug Then
        g_WordAp.Visible = True
        g_WordAp.Activate
      End If
      'end 2025/3/19
      
      'Add By Sindy 2013/11/14 +因加右代表圖,所以調整其他欄位的位置
      'Add By Sindy 2014/3/5 P,CFP也要加右代表圖
      'Modified by Lydia 2018/08/22 所有 T字頭的系統類別, 都加代表圖
      'If Text1 = "T" Or Text1 = "CFT" Or Text1 = "P" Or Text1 = "CFP" Then
      If Left(Text1, 1) = "T" Or Text1 = "CFT" Or Text1 = "P" Or Text1 = "CFP" Then
         '帶出系統日期
         'modify by sonia 2019/11/15 改以民國(西元)年月日表示,並依iPicNo調整位置
         '.Selection.TypeText "　　　　　　　　　　　　　　　　　　　　中華民國" & Val(Mid(strSrvDate(1), 1, 4)) - 1911 & "(" & Mid(strSrvDate(1), 1, 4) & ")年" & Mid(strSrvDate(1), 5, 2) & "月   日"
         If iPicNo = 10 Then
            'Modified by Morgan 2022/2/21 改用函數以確保與更新發文日函數格式一致
            '.Selection.TypeText "　　　　　　　　　　　　　　　　　　　 中華民國" & Val(Mid(strSrvDate(1), 1, 4)) - 1911 & "(" & Mid(strSrvDate(1), 1, 4) & ")年" & Mid(strSrvDate(1), 5, 2) & "月   日"
            .Selection.TypeText "　　　　　　　　　　　　　　　　　　　 " & PUB_GetCusLetterDate
         Else
            'Modified by Morgan 2022/2/21 改用函數以確保與更新發文日函數格式一致
            '.Selection.TypeText "　　　　　　　　　　　　　　　　　　　　中華民國" & Val(Mid(strSrvDate(1), 1, 4)) - 1911 & "(" & Mid(strSrvDate(1), 1, 4) & ")年" & Mid(strSrvDate(1), 5, 2) & "月   日"
            .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　" & PUB_GetCusLetterDate
         End If
         'end 2019/11/15
         .Selection.TypeParagraph
      End If
      '2013/11/14 END
      
      intLenText = 0 'Add By Sindy 2013/11/14
      '郵遞區號
      .Selection.TypeText stZip
      intLenText = intLenText + Len(stZip) 'Add By Sindy 2013/11/14
            
      'Modified by Morgan 2021/12/2 +bolECust
      strLetterNo = GetLetterNo(stCP09, bolECust) 'Added by Morgan 2014/12/8
      'end 2021/12/2
      
      'Added by Morgan 2015/11/12
      '商標案點選的收文號有下一程序的要掛號
      strExc(2) = ""
      'Modified by Morgan 2019/2/14 沒有定稿的也要判斷下一程序 Ex:T-209730 (108/1/29 其他來函)
      'If Combo8.ListIndex > 0 Then
      '   If InStr(text1, "T") > 0 And Combo8.ItemData(Combo8.ListIndex) <> 0 Then
      If InStr(Text1, "T") > 0 Then
         strExc(1) = ""
         If Combo8.ItemData(Combo8.ListIndex) <> 0 Then
            strExc(1) = PUB_Num2DocNo(Combo8.ItemData(Combo8.ListIndex))
         ElseIf stCP09 <> "" Then
            strExc(1) = stCP09
         End If
         If strExc(1) <> "" Then
      'end 2019/2/14
            'Add By Sindy 2025/8/12
            If m_T727RecvNo <> "" Then
               strLetterNo = Right(m_T727RecvNo, 6)
               strExc(2) = "Y" '要掛號
               If m_T727CP43No <> "" Then
                  '另相關總收文號來函性質發文日若為11/11/11，則分析撰寫信函必須帶掛號，否則不必加掛號。
                  strExc(0) = "Select CP27 FROM CaseProgress" & _
                              " WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP09='" & m_T727CP43No & "'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If Val("" & RsTemp.Fields(0)) > 0 And Val("" & RsTemp.Fields(0)) <> 19221111 Then
                        strExc(2) = "" '相關總收文號已有發文日,不必加掛號
                     End If
                  End If
               End If
            Else
            '2025/8/12 END
               strExc(0) = "select np09 from CaseProgress,nextprogress where cp09='" & strExc(1) & "' and cp07>0 and np01(+)=cp09 and np06 is null and np09>0"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strExc(2) = "Y"
               End If
            End If
         End If
      End If
      
      'Added by Morgan 2021/11/18
      'Modify By Sindy 2025/8/12 + Or (m_T727RecvNo <> "" And strExc(2) = ""): 不印掛號
      If lngPaper = vbNo Or (m_T727RecvNo <> "" And strExc(2) = "") Then
         '全E化不印掛號
      'end 2021/11/18
      ElseIf strExc(2) = "Y" Then
         .Selection.TypeText String(14 - Len(stZip), "　")
         .Selection.Font.Borders(1).LineStyle = .Options.DefaultBorderLineStyle
         'Modified by Morgan 2021/10/12 X15416020(宋健民)案件掛號改為雙掛號
         If m_CustNo(1) = "X15416020" Then
            .Selection.TypeText "雙掛號"
            intLenText = 14 + Len("雙掛號")
         Else
            .Selection.TypeText "掛號"
            intLenText = 14 + Len("掛號")
         End If
         'end 2021/10/12
         .Selection.Font.Borders(1).LineStyle = wdLineStyleNone
      Else
      'end 2015/11/12
      
         'Add by Morgan 2008/8/8 專利處+掛號,內商改在下面控制
         If m_StrUserST03 = "P10" Or m_StrUserST03 = "P11" Or m_StrUserST03 = "P14" Then
            'Modified by Morgan 2012/1/2 檢索報告14不掛號--游登銘
            'Modified by Morgan 2013/2/20 控制P案就好--游登銘
            'If m_Combo8 <> 14 Then
            'Modified by Morgan 2014/10/8 + 專利權評價報告 "17" --玲玲
            'Modified by Morgan 2022/9/2 分析及一般信函也不掛號--郭雅娟
            'Modified by Morgan 2024/5/8 有回覆單的也要掛號 Ex:P131131(CB3022932)
            If Not (Text1 = "P" And (m_Combo8 = "14" Or m_Combo8 = "17")) And Not (Trim(Right(Combo8, 4)) = "941" Or Combo8 = "一般格式") Or strReturnSheet <> "" Then
               .Selection.TypeText String(14 - Len(stZip), "　")
               .Selection.Font.Borders(1).LineStyle = .Options.DefaultBorderLineStyle
               'Modified by Morgan 2021/10/12 X15416020(宋健民)案件掛號改為雙掛號
               If m_CustNo(1) = "X15416020" Then
                  .Selection.TypeText "雙掛號"
                  intLenText = 14 + Len("雙掛號")
               Else
                  .Selection.TypeText "掛號"
                  intLenText = 14 + Len("掛號") 'Add By Sindy 2013/11/14
               End If
               'end 2021/10/12
               .Selection.Font.Borders(1).LineStyle = wdLineStyleNone
            End If
         End If
         'end 2008/8/8
      
         '2008/9/22 add by sonia 商標處核駁1002,核駁前先行通知1202,敗訴1004,部分核駁1205,其他有期限之來函要加印掛號
         Select Case Left(m_Combo8, 1)
            Case "2", "4", "5", "6", "8"
               '2012/4/27 MODIFY BY SONIA 加6B
               'MODIFY BY SONIA 2015/9/16 加42A大陸異議敗訴無期限
               'modify by sonia 2018/9/26 取消42A,改印掛號-林純貞
               If m_Combo8 <> "64" And m_Combo8 <> "68" And m_Combo8 <> "6B" Then     '64,68,6B無期限
                  .Selection.TypeText String(14 - Len(stZip), "　")
                  .Selection.Font.Borders(1).LineStyle = .Options.DefaultBorderLineStyle
                  'Modified by Morgan 2021/10/12 X15416020(宋健民)案件掛號改為雙掛號
                  If m_CustNo(1) = "X15416020" Then
                     .Selection.TypeText "雙掛號"
                     intLenText = 14 + Len("雙掛號")
                  Else
                     .Selection.TypeText "掛號"
                     intLenText = 14 + Len("掛號") 'Add By Sindy 2013/11/14
                  End If
                  'end 2021/10/12
                  .Selection.Font.Borders(1).LineStyle = wdLineStyleNone
'2018/9/28 cancel by sonia 還是先給智權人員
'               'add by sonia 2018/9/26 C類無期限改印平信(發文室要直寄) 64被異議,被評定,被廢止 68智慧局答辯函 6B對方撤回
'               Else
'                  .Selection.TypeText String(14 - Len(stZip), "　")
'                  .Selection.TypeText "平信"
'                  intLenText = 14 + Len("平信")
'                  .Selection.Font.Borders(1).LineStyle = wdLineStyleNone
'               'end 2018/9/26
'end 2018/9/28
               End If
         End Select
         '2008/9/22 end
      
         '2008/10/20 其他部門承辦人使用時加掛號
         If m_StrUserST03 = "F10" Or m_StrUserST03 = "F11" Or m_StrUserST03 = "F21" Or m_StrUserST03 = "F81" Or m_StrUserST03 = "L01" Then
            .Selection.TypeText String(14 - Len(stZip), "　")
            'Modified by Morgan 2021/10/12 X15416020(宋健民)案件掛號改為雙掛號
            If m_CustNo(1) = "X15416020" Then
               .Selection.TypeText "雙掛號"
               intLenText = 14 + Len("雙掛號")
            Else
               .Selection.Font.Borders(1).LineStyle = .Options.DefaultBorderLineStyle
               .Selection.TypeText "掛號"
               intLenText = 14 + Len("掛號") 'Add By Sindy 2013/11/14
               .Selection.Font.Borders(1).LineStyle = wdLineStyleNone
            End If
            'end 2021/10/12
         End If
         '2008/10/20 end
      
      End If 'Added 2015/11/12
      
       
      'Add By Sindy 2013/11/14 +因加右代表圖,所以調整其他欄位的位置
      'Add By Sindy 2014/3/5 P,CFP也要加右代表圖
      'Modified by Morgan 2014/12/8 發文字號改抓變數
      'Modified by Lydia 2018/08/22 所有 T字頭的系統類別, 都加代表圖
      'If Text1 = "T" Or Text1 = "CFT" Or Text1 = "P" Or Text1 = "CFP" Then
      If Left(Text1, 1) = "T" Or Text1 = "CFT" Or Text1 = "P" Or Text1 = "CFP" Then
         '中文也要抓專業代號
         strExc(0) = "SELECT ST07 FROM STAFF WHERE ST01='" & strUserNum & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strText = "(" & Val(Mid(strSrvDate(1), 1, 4)) - 1911 & ")晉" & RsTemp.Fields(0) & "字第" & strLetterNo & "號"
         Else
            strText = "(" & Val(Mid(strSrvDate(1), 1, 4)) - 1911 & ")晉字第" & strLetterNo & "號"
         End If
         'modify by sonia 2019/11/18 iPicNo=10者配合發文字號提前一格半形
         '.Selection.TypeText String(20 - intLenText, "　")
         If iPicNo = 10 Then
            .Selection.TypeText String(19 - intLenText, "　") & " "
         Else
            .Selection.TypeText String(20 - intLenText, "　")
         End If
         'end 2019/11/18
         .Selection.TypeText strText
      End If
      '2013/11/14 END
      .Selection.TypeParagraph
      iLineCount = iLineCount + 1

      'Add By Sindy 2013/11/14 +右代表圖
      'Add By Sindy 2014/3/5 P,CFP也要加右代表圖
      'Modified by Lydia 2018/08/22 所有 T字頭的系統類別, 都加代表圖
      'If Text1 = "T" Or Text1 = "CFT" Or Text1 = "P" Or Text1 = "CFP" Then
      If Left(Text1, 1) = "T" Or Text1 = "CFT" Or Text1 = "P" Or Text1 = "CFP" Then
         .Selection.TypeText "|#右代表圖#|"
'         m_MySt(1) = Text1
'         m_MySt(2) = Text2
'         m_MySt(3) = Text3
'         m_MySt(4) = Text4
         m_bolNewFormat = True
         PUB_AddInPicToWordR g_WordAp
         m_bolNewFormat = False '恢復預設值
      End If
      '2013/11/14 End
      
      '地址
      If stAddr <> "" Then
         strExc(0) = stAddr
         Do While Len(strExc(0)) > 17
            .Selection.TypeText Left(strExc(0), 17)
            .Selection.TypeParagraph
            strExc(0) = Mid(strExc(0), 18)
            iLineCount = iLineCount + 1
         Loop
         If strExc(0) <> "" Then
            .Selection.TypeText strExc(0)
            .Selection.TypeParagraph
            iLineCount = iLineCount + 1
         End If
      End If
      '收件人
      If stReceiver <> "" Then
         strExc(0) = stReceiver
         Do While Len(strExc(0)) > 17
            .Selection.TypeText Left(strExc(0), 17)
            .Selection.TypeParagraph
            strExc(0) = Mid(strExc(0), 18)
            iLineCount = iLineCount + 1
         Loop
         .Selection.TypeText strExc(0)
         .Selection.TypeParagraph
         iLineCount = iLineCount + 1
      End If
      If GetTextLength(stContact) < 12 Then
         .Selection.TypeText stContact & String(12 - GetTextLength(stContact), " ") & "鈞啟(" & CaseNo & ")"
      Else
         .Selection.TypeText stContact & " 鈞啟(" & CaseNo & ")"
      End If
      
      .Selection.TypeParagraph
      iLineCount = iLineCount + 1
      
      'Added by Lydia 2020/08/26 (限外專)受文者非X編號，不論是FC或CF代理人以及語言都帶備註出來；其他部門的都先不動，將來有需要再加。
      If strSrvDate(1) >= 各項指示啟用日 And Left(m_StrUserST03, 2) = "F2" And (Option2.Value = True Or Option6.Value = True) And Text5 & Text6(0) & Text6(1) <> "" Then
           .Selection.TypeParagraph
           PrintMemo
           .Selection.TypeParagraph
      Else
      'end 2020/08/26
            '補滿5行
            For intI = iLineCount + 1 To 5
               .Selection.TypeParagraph
            Next
      End If 'Added by Lydia 2020/08/26
      
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      
      'Modified by Morgan 2016/11/11 stAppName
      '.Selection.TypeText "致：" & stReceiver & stApp1Title
      'Modified by Morgan 2025/4/9
      '.Selection.TypeText "致：" & stAppName & stApp1Title
      .Selection.TypeText strData
      'end 2025/4/9
      'end 2016/11/11
      .Selection.TypeParagraph
      
      'Add by Morgan 2007/1/22 傳真號碼
      strExc(0) = m_strFax(0)
      If m_strFax(2) <> "" Then
         If m_strFax(0) <> "" Then
            strExc(0) = strExc(0) & "," & m_strFax(2)
         Else
            strExc(0) = m_strFax(2)
         End If
      End If
      
      strExc(1) = m_strFax(1)
      If m_strFax(3) <> "" Then
         If m_strFax(1) <> "" Then
            strExc(1) = strExc(1) & "," & m_strFax(3)
         Else
            strExc(1) = m_strFax(3)
         End If
      End If
      
      .Selection.Font.Size = 10 'Add By Sindy 2014/3/13 因右代表圖之故會蓋掉FAX資料,因此縮小字的size
      '2008/7/23 modify by sonia 專利處人員取消不印
      '.Selection.TypeText "　　TEL：" & strExc(1) & "　　FAX：" & strExc(0)
      If Mid(m_Dept, 1, 2) <> "P1" Then
         .Selection.TypeText "TEL：" & strExc(1) & "　FAX：" & strExc(0)
      End If
      '2008/7/23 end
      'end 2007/1/22
      .Selection.Font.Size = 14 'Add By Sindy 2014/3/13 還原字的size
      
      .Selection.TypeParagraph
      
      'Added by Morgan 2016/11/11
      If pNoCopy = False And stCC <> "" Then
         If pIsCopy Then
            .Selection.TypeText "|#(框線)副本#|：" & stCC
         Else
            .Selection.TypeText "副本：" & stCC
         End If
         .Selection.TypeParagraph
      End If
      
      'Modify By Sindy 2013/11/14 +if
      'Add By Sindy 2014/3/5 P,CFP也要加右代表圖
      'Modified by Lydia 2018/08/22 所有 T字頭的系統類別, 都加代表圖
      'If Text1 <> "T" And Text1 <> "CFT" And Text1 <> "P" And Text1 <> "CFP" Then
      If Left(Text1, 1) <> "T" And Text1 <> "CFT" And Text1 <> "P" And Text1 <> "CFP" Then
         '帶出系統日期
         '2008/6/26 MODIFY BY SONIA 往前移一格,否則 專六十六 會跑到下一行
         '2008/7/30 MODIFY BY SONIA 只帶年月不帶日
         '.Selection.TypeText "　　　　　　　　　　　　　　　　　　　　中華民國 " & Mid(strSrvDate(2), 1, 2) & "年" & Mid(strSrvDate(2), 3, 2) & "月" & Mid(strSrvDate(2), 5, 2) & "日"
         'Modified by Morgan 2013/7/5
         '.Selection.TypeText "　　　　　　　　　　　　　　　　　　　　中華民國 " & Val(Mid(strSrvDate(1), 1, 4)) - 1911 & "年" & Mid(strSrvDate(1), 5, 2) & "月    日"
         'modify by sonia 2019/11/15 改以民國(西元)年月日表示
         .Selection.TypeText "　　" & stMultiApp & "　　　　　　　　　　　　中華民國" & Val(Mid(strSrvDate(1), 1, 4)) - 1911 & "(" & Mid(strSrvDate(1), 1, 4) & ")年" & Mid(strSrvDate(1), 5, 2) & "月   日"      'end 2013/7/5
         'end 2013/7/5
         .Selection.TypeParagraph
         
         'Modify by Morgan 2005/9/19 中文也要抓專業代號
         '.Selection.TypeText "　　　　　　　　　　　(" & Mid(strSrvDate(2), 1, 2) & ")晉字第　　　　　號"
         strExc(0) = "SELECT ST07 FROM STAFF WHERE ST01='" & strUserNum & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            '2008/6/26 MODIFY BY SONIA 往前移一格,否則 專六十六 會跑到下一行
            .Selection.TypeText "敬啟者：　　　　　　　　　　　　　　　　(" & Val(Mid(strSrvDate(1), 1, 4)) - 1911 & ")晉" & RsTemp.Fields(0) & "字第　　　　號"
         Else
            .Selection.TypeText "敬啟者：　　　　　　　　　　　　　　　　(" & Val(Mid(strSrvDate(1), 1, 4)) - 1911 & ")晉字第　　　　號"
         End If
      Else
         .Selection.TypeText "　　" & stMultiApp
         .Selection.TypeParagraph
         .Selection.TypeText "敬啟者："
      End If
      '2013/11/14 END
      
      '2008/5/1 add by sonia 專利處要求加印
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      
      Call WordChinese_sub1(g_WordAp)   '2019/11/20 將信函內文拆至WordChinese_sub1,WordChinese_sub2
      
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeText "　　耑此　　順頌"
      .Selection.TypeParagraph
      'Modify By Sindy 2012/6/6
      '.Selection.TypeText "商祺"
      .Selection.TypeText "時祺"
      '2012/6/6 End
      .Selection.TypeParagraph
      '2008/5/1 end
      .Selection.TypeParagraph
      '2008/7/23 MODIFY BY SONIA 專利處取消 備註
      'If Combo8.Text = "一般格式" Then
      If m_Combo8 = "00" And Mid(m_Dept, 1, 2) <> "P1" Then
      '2008/7/23 END
         .Selection.TypeText "備註："
         .Selection.TypeParagraph
      End If
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeParagraph
'2015/7/15 modify by sonia 定稿已全加智權人員,故此處也全部都加,另公司落款及智權人員都改為右靠
      'Modify by Morgan 2007/10/2 --秀玲
      '   .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　台一國際商標專利事務所  敬上"
      '商標
      
      'Modified by Morgan 2012/9/5 加專利案以專利商標出名者
      'Modified by Morgan 2014/1/17 +J公司判斷改和信頭一樣
      'If CheckSys(pa(1)) = "2" Or iPicNo = 12 Then
      '   .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　台一國際專利商標事務所  敬上"
      ''其他
      'Else
      '   .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　台一國際專利法律事務所  敬上"
      'End If
      ''end 2007/10/2
'      '專利商標
'      If m_CompNo = "T" Then
'         .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　台一國際專利商標事務所  敬上"
'      '智權公司
'      ElseIf m_CompNo = "J" Then
'         .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　台一智權股份有限公司  敬上"
'      '未設定
'      Else
'         '專利商標
'         'modify by sonia 2014/4/28
'         'If Text1.Text = "T" Then
'         If InStr(Text1.Text, "T") > 0 Then
'            .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　台一國際專利商標事務所  敬上"
'         '專利法律
'         Else
'            .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　台一國際專利法律事務所  敬上"
'         End If
'      End If
'      'end 2014/1/17
'      .Selection.TypeParagraph
      
'2014/3/14 CANCEL BY SONIA
'      'Add By Sindy 2013/2/8 +代表圖
'      '若客戶為順德X07166及其關係企業,且為專利案件P, PS, CFP, CPS則加代表圖
'      If Left(Trim(m_CustNo(1)), 6) = "X07166" And _
'         (Text1 = "P" Or Text1 = "PS" Or Text1 = "CFP" Or Text1 = "CPS") Then
'         .Selection.TypeText "|#代表圖#|"
''         m_MySt(1) = Text1
''         m_MySt(2) = Text2
''         m_MySt(3) = Text3
''         m_MySt(4) = Text4
'         Call PUB_AddInPicToWord(g_WordAp, 34)
'      End If
'      '2013/2/8 End
      
      '2008/5/1 add by sonia 專利處要求加印智權人員
      '2008/7/23 modify by sonia 程序才要加印
      '2008/9/18 modify by sonia 內外商也加
'      If m_Dept = "P12" Or Left(m_Dept, 2) = "F1" Or Left(m_Dept, 2) = "P2" Then
'         '2011/8/5 ADD BY SONIA
'         If custarea = "" Then
'            .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　　　" & custsales
'            .Selection.TypeParagraph
'         Else
'         '2011/8/5 END
'            'Modified by Morgan 2015/4/16 有跳行問題，改右靠
'            '.Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　　　" & custarea & "　" & custsales
'            '.Selection.TypeParagraph
'            .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
'            .Selection.TypeText custarea & "　" & custsales
'            .Selection.TypeParagraph
'            .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
'            'end 2015/4/16
'         End If   '2011/8/5 ADD BY SONIA
'      'End If
      '2008/5/1 end
      
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      

'Added by Morgan 2020/3/25
If strSrvDate(1) >= 智慧所更名日 Then
   'Modified by Morgan 2021/8/17 +e化客戶加印公司章
   .Selection.TypeText CompNameQuery(m_CompNo) & IIf(lngPaper = 0, "", "|#(自動公司章)#|") & "  敬上"
Else
'end 2020/3/25

      '專利商標
      If m_CompNo = "T" Then
         .Selection.TypeText "台一國際專利商標事務所  敬上"
      '智權公司
      ElseIf m_CompNo = "J" Then
         .Selection.TypeText "台一智權股份有限公司  敬上"
      '未設定
      Else
         '專利商標
         If InStr(Text1.Text, "T") > 0 Then
            .Selection.TypeText "台一國際專利商標事務所  敬上"
         '專利法律
         Else
            .Selection.TypeText "台一國際專利法律事務所  敬上"
         End If
      End If
      
End If 'Added by Morgan 2020/3/25

      .Selection.TypeParagraph
         
      If custarea = "" Then
         .Selection.TypeText "" & custsales
         .Selection.TypeParagraph
      Else
         .Selection.TypeText custarea & "　" & custsales
         .Selection.TypeParagraph
      End If
      'Added by Lyda 2020/07/16 法律所案源收文：撰寫信函落款帶最新A類承辦，同時加帶聯絡人(案源介紹人)
      If stLos04 <> "" Then
          If InStr(stLos04, ",") > 0 Then
              stLos04 = Mid(1, InStr(stLos04, ",") - 1)
              stLos04Name = Mid(1, InStr(stLos04Name, ",") - 1)
          End If
         .Selection.TypeText "聯絡人　" & GetDepartmentName(GetSalesArea(stLos04)) & "　" & stLos04Name
         .Selection.TypeParagraph
      End If
      
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
'2015/7/15 END
      
      '2008/9/18 ADD BY SONIA 內外商加印分所案號
      'Modify by Morgan 2010/6/15 因CFP也要，故+程序
      If (m_Dept = "P12" Or Left(m_Dept, 2) = "F1" Or Left(m_Dept, 2) = "P2") And casetype5 <> "" Then
         .Selection.TypeText "　　　　      　　　　　　　　　　　　　　　　（分所案號：" & casetype5 & "）"
         .Selection.TypeParagraph
      End If
      '2008/5/1 end
       
      '2008/6/25 MODIFY BY SONIA 最終核駁,通知要求選取 還有附錄
      Select Case m_Combo8
         Case "12"      'CFP最終核駁
            .Selection.TypeText Chr(12)
            .Selection.TypeText "附錄："
            .Selection.TypeParagraph
            'Modified by Morgan 2015/5/6
            '.Selection.TypeText "　　美國專利法對於最終審定後進行各項程序及所需之延期費規定如下："
            '.Selection.TypeParagraph
            '.Selection.TypeText "　　一、若審查員在最終審定發出的二個月內收到申請人之答辯，則審查員有義務做出建議性之處分。"
            '.Selection.TypeParagraph
            '.Selection.TypeText "　　若申請人在收到審查員的建議性處分後提起訴願或請求繼續審查，而提起訴願或請求繼續審查的日期又晚於最終審定發出的三個月，則申請人必須在提起訴願或請求繼續審查之同時繳交延期費。延期費之起算日係審查員做出建議性處分之日與審查員做出最終審定起三個月之日二者中之較晚者。"
            '.Selection.TypeParagraph
            '.Selection.TypeText "　　二、若審查員在最終審定發出的二個月至三個月之間收到申請人之答辯，則審查員亦有義務做出建議性之處分。"
            '.Selection.TypeParagraph
            '.Selection.TypeText "　　若申請人在收到審查員的建議性處分後提起訴願或請求繼續審查，而提起訴願或請求繼續審查的日期又晚於最終審定發出的三個月(通常皆晚於三個月) ，則申請人必須在提起訴願或請求繼續審查之同時繳交延期費。延期費之起算日係審查員做出最終審定之日起三個月之日。"
            '.Selection.TypeParagraph
            '.Selection.TypeText "　　三、若審查員在三個月以後才收到申請人之答辯，則審查員在研究申請人的答辯之後並無義務做出建議性之處分。此時申請人必須於最終審定發出之六個月內自行決定是否提起訴願或請求繼續審查，否則申請案審查程序便告終止，即為審查確定。"
            '.Selection.TypeParagraph
            '.Selection.TypeText "　　若申請人提起訴願或請求繼續審查(必然晚於三個月)，則申請人必須在提起訴願或請求繼續審查之同時繳交延期費。延期費之起算日係審查員做出最終審定之日起三個月之日。簡言之，若申請人在審查員做出最終審定之日起三個月之後才進行非正式程序(答辯)與正式程序(訴願或請求繼續審查)，則必須繳交二次延期費。"
            .Selection.TypeText "　　在最終審定發出後，申請人可以選擇提出請求繼續審查、答辯或是訴願。而前兩者是較為常見的程序。以下為最終審定後的相關程序資訊，供申請人參考。"
            .Selection.TypeParagraph
            .Selection.TypeParagraph
            .Selection.Font.Bold = True
            .Selection.TypeText "　　一、最終審定後提出答辯及修正限制"
            .Selection.Font.Bold = False
            .Selection.TypeParagraph
            .Selection.TypeText "　　美國專利法規定，在最終審定發出後，若申請人提出答辯，則僅得在以下兩種情況中對請求項、說明書或圖式進行修正："
            .Selection.TypeParagraph
            .Selection.TypeText "　　（一）、刪除請求項；或"
            .Selection.TypeParagraph
            .Selection.TypeText "　　（二）、根據前一次審查意見通知書中的形式要求進行修改請求項、說明書或圖式。"
            .Selection.TypeParagraph
            .Selection.TypeText "　　根據上述規定，在接獲最終審定後，若申請人提出答辯，則在答辯時任何對請求項、說明書或圖式的實質修正，均不會被美國專利局所接受。因此，在接獲最終審定後，若申請人欲對請求項進行實質修正，例如將說明書或圖式所揭露內容納入申請專利範圍而導致修正後請求項屬前所未見，則申請人必須提出請求繼續審查，而非提出答辯。"
            .Selection.TypeParagraph
            .Selection.TypeText "　　申請人應注意，若晚於最終審定發出的三個月提出答辯，審查委員將不具回覆與處分之義務，故申請人應最遲於最終審定後三個月以內提出答辯。"
            .Selection.TypeParagraph
            .Selection.TypeParagraph
            .Selection.Font.Bold = True
            .Selection.TypeText "　　二、最終審定後提出請求繼續審查"
            .Selection.Font.Bold = False
            .Selection.TypeParagraph
            .Selection.TypeText "　　若申請人於最終審定發出後提出請求繼續審查，則重啟本案審查程序，申請人可於提出請求繼續審查時，可在不超出原申請時揭露內容的前提下，一併對本案請求項、說明書或圖式進行任何實質修正。"
            .Selection.TypeParagraph
            .Selection.TypeText "　　此外，若晚於最終審定發出的三個月提出請求繼續審查，則須繳交額外延期費。"
            .Selection.TypeParagraph
            .Selection.TypeParagraph
            .Selection.Font.Bold = True
            .Selection.TypeText "　　三、最終審定後相關程序之延期費計算"
            .Selection.Font.Bold = False
            .Selection.TypeParagraph
            .Selection.TypeText "　　美國專利法對於最終審定後進行各項程序及所需之延期費規定如下："
            .Selection.TypeParagraph
            .Selection.TypeText "　　（一）、若審查員在最終審定發出的二個月內收到申請人之答辯，則審查員有義務做出建議性之處分。"
            .Selection.TypeParagraph
            .Selection.TypeText "　　若申請人在收到審查員的建議性處分後提出訴願或請求繼續審查，而提出訴願或請求繼續審查的日期又晚於最終審定發出的三個月，則申請人必須在提出訴願或請求繼續審查之同時繳交延期費。延期費之起算日係審查員做出建議性處分之日與審查員做出最終審定起三個月之日二者中之較晚者。"
            .Selection.TypeParagraph
            .Selection.TypeText "　　（二）、若審查員在最終審定發出的二個月至三個月之間收到申請人之答辯，則審查員亦有義務做出建議性之處分。"
            .Selection.TypeParagraph
            .Selection.TypeText "　　若申請人在收到審查員的建議性處分後提出訴願或請求繼續審查，而提出訴願或請求繼續審查的日期又晚於最終審定發出的三個月(通常皆晚於三個月) ，則申請人必須在提出訴願或請求繼續審查之同時繳交延期費。延期費之起算日係審查員做出最終審定之日起三個月之日。"
            .Selection.TypeParagraph
            .Selection.TypeText "　　（三）、若審查員在三個月以後才收到申請人之答辯，則審查員在研究申請人的答辯之後並無義務做出建議性之處分。此時申請人必須於最終審定發出之六個月內自行決定是否提出訴願或請求繼續審查，否則申請案審查程序便告終止，即為審查確定。"
            .Selection.TypeParagraph
            .Selection.TypeText "　　（四）、若申請人在收到上述建議性處分後提出訴願或請求繼續審查（必然晚於三個月），則申請人必須在提出訴願或請求繼續審查之同時繳交延期費。延期費之起算日係審查員做出最終審定之日起三個月之日。簡言之，若申請人在審查員做出最終審定之日起三個月之後才進行非正式程序（答辯）與正式程序（訴願或請求繼續審查），則必須繳交二次延期費。"
            'end 2015/5/6
         Case "13"      'CFP通知要求選取
            .Selection.TypeText Chr(12)
            .Selection.TypeText "附錄："
            .Selection.TypeParagraph
            .Selection.TypeText "　　美國有關選取的相關規定則請參閱下列的說明："
            .Selection.TypeParagraph
            'Modify By Sindy 2012/8/29 修改定稿內容
'            .Selection.TypeText "　　‧在選取經審查後，假如本案具有一項具有可專利性且包含各發明的概括性申請專利範圍時，未被選取的部分亦會被考量其可專利性；"
'            .Selection.TypeParagraph
'            .Selection.TypeText "　　‧所選取的部分將會持續地進行審查；"
'            .Selection.TypeParagraph
'            .Selection.TypeText "　　未選取的部分在母案仍處於審查階段(pending)而提出分割的請求後，其會擁有一個新的申請日以及申請序號，但其新穎性的考量則是以母案的申請日起算；"
'            .Selection.TypeParagraph
'            .Selection.TypeText "　　‧分割案的專利期限是以母案的申請日起算二十年(發明案)；"
'            .Selection.TypeParagraph
'            .Selection.TypeText "　　‧未選取的部分在母案處於核准但未公告時，所提出之分割請求仍擁有母案之較早申請日之優惠；"
'            .Selection.TypeParagraph
'            .Selection.TypeText "　　‧未選取的部分在母案處於核准且公告時，才提出分割請求的話，其新穎性將由分割案送件日開始考慮；此時，將會有下列兩種狀況："
'            .Selection.TypeParagraph
'            .Selection.TypeText "　　　1.母案被引證為前案，而對分割案作核駁的審定；"
'            .Selection.TypeParagraph
'            .Selection.TypeText "　　　2.因為重複專利，會造成分割案必須要提出放棄部分專利之聲明。"
'            .Selection.TypeParagraph
'            .Selection.TypeText "　　然而，在提出分割案時，　" & custtype & "除需注意上述因為提出時機所造成的後果不同之外，尚有一件與提出時機極為相關的事項-前案揭露聲明(INFORMATION DISCLOSURE STATEMENT, IDS)"
'            .Selection.TypeParagraph
'            .Selection.TypeText "　　‧新案提申時若沒有提出IDS，則在接或要求作選取之審定書的同時提出分割案的話，因為母案並未進行任何檢索的動作，故而此時分割案的提出並不會被要求提出IDS，但　" & custtype & "仍可主動地對美國專利局提出IDS之呈報；"
'            .Selection.TypeParagraph
'            .Selection.TypeText "　　‧若是在接獲第一次核駁通之後，才提出分割案的話，則因為母案內已含有審查委員所引證之前案，故而此一階段所提出之分割案必定會被要求提出IDS；此時若　" & custtype & "決定不提的話，國外代理人亦會因為美國專利法之規定主動地為　" & custtype & "提出IDS的資料。"
            .Selection.TypeText "　　‧所選取的部分將會進行審查，未選被選取的部分則將不被審查，但在審查後，假如原案具有一項具有可專利性且包含各發明的概括性申請專利範圍時，未被選取的部分亦會被考量其可專利性。"
            .Selection.TypeParagraph
            .Selection.TypeText "　　‧未被選取的部分在原案審查階段或在原案核准公告前可提出分割申請，此分割案雖會賦予新的申請日及申請案號，但該分割案可專利性的考量則是以原案的申請日為基準日。"
            .Selection.TypeParagraph
            .Selection.TypeText "　　‧在原案核准公告後方欲提出分割申請案，則無法獲得原案較早申請日的利益。"
            .Selection.TypeParagraph
            'Removed by Morgan 2022/1/12 是錯誤的--郭
            '.Selection.TypeText "　　‧若在原案接獲審查意見書後才提出分割申請案，分割案應就原案審查意見書所檢附之引證前案，提出前案揭露聲明(INFORMATION DISCLOSURE STATEMENT, IDS)。"
            '.Selection.TypeParagraph
            'end 2022/1/12
            'Added by Morgan 2020/2/11 不論是否有主張都帶--郭
            .Selection.TypeText "　　‧原案若有主張優先權，分割申請案仍應主張優先權。"
            .Selection.TypeParagraph
            'end 2020/2/11
            'Modified by Morgan 2018/9/28
            '.Selection.TypeText "　　‧美國發明分割案的專利權期限是以原案的申請日起算二十年，設計分割案的專利權期限則係由該分割案本身公告日起十四年。"
            .Selection.TypeText "　　‧美國發明分割案的專利權期限是以原案的申請日起算二十年，設計分割案的專利權期限則係由該分割案本身公告日起十五年。"
            'end 2018/9/28
      End Select
      '2008/6/25 end
      
     'Added by Morgan 2023/6/21
     If strReturnSheet <> "" Then
         .Selection.TypeText Chr(12)
         .Selection.TypeText strReturnSheet
     End If
     'end 2023/6/21
     
     '2008/7/23 add by sonia 英數字改字體
      If Mid(m_Dept, 1, 2) = "P1" Then
         .Selection.WholeStory
         'modify by sonia 2019/11/22 郭同意改Arial為Times New Roman(與定稿統一),否則發文字號位置會跑掉
         .Selection.Font.Name = "Times New Roman"
      End If
     '2008/7/23 end
     
     'Added by Morgan 2016/11/11
     .Selection.WholeStory
      ChgWordFormat g_WordAp, .Selection.Text
      'end 2016/11/11
      
   End With
   '2008/9/19 ADD BY SONIA
   Select Case Left(m_Combo8, 1)
      Case "2", "3", "4", "5", "6", "8"
         PhaseIndent    '調整首行凸排
   End Select
   '2008/9/19 END
      
      
  g_WordAp.Visible = True
  g_WordAp.WindowState = wdWindowStateMaximize
  Set g_WordAp = Nothing 'Added by Morgan 2015/9/7

ERRORSECTION1:
   If Err.NUMBER <> 0 Then
      Select Case Err.NUMBER
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

'Add by Morgan 2006/10/2 大陸信函格式
'Modify by Morgan 2008/7/17 改開窗信封用的信紙格式
Private Sub WordChinese1()
'Add by Morgan 2008/7/17
Dim stReceiver As String '收件人
Dim stAddr As String '地址
Dim stContact As String '接洽人
Dim iLineCount As Integer '行數
'Add By Sindy 2009/10/15
Dim strText As String
Dim m_CP07 As String '法定期限/西元
Dim m_CP06 As String '本所期限/西元
Dim m_DueDT As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'2009/10/15 End
Dim iPicNo As Integer, stFileName As String 'Add by Morgan 2010/12/13
Dim iPicNo2 As Integer 'Added by Morgan 2012/12/27
Dim oShape 'Added by Morgan 2011/12/13
Dim m_DueDate As String 'Add By Sindy 2012/4/23
Dim strSignFile As String, BolHaveImg As Boolean 'Added by Morgan 2014/11/14
'Added by Lydia 2017/05/03
Dim iPicPos As Double '圖片的位置
Dim pageY As Double '圖片的上邊界
Dim bolP020 As Boolean '是否為P大陸案
Dim bVisibled As Boolean 'Added by Lydia 2017/05/18 記錄Word是否顯示
Dim bolBCP As Boolean 'Added by Morgan 2023/9/12 是否有B類未發文

'Modified by Morgan 2020/3/25
If strSrvDate(1) >= 智慧所更名日 Then
   m_CompNo = PUB_GetReceiptComp(pa(1), pa(2), pa(3), pa(4))
   PUB_GetLetterPicID m_CompNo, pa(1), iPicNo, iPicNo2, 1, True, m_Dept
Else
   m_CompNo = ""

   iPicNo = 7 'Add by Morgan 2010/12/13 加可印信頭

   'Added by Morgan 2013/12/27
   'Modified by Morgan 2014/6/27
   'If Text1.Text = "P" Then
   If Left(m_Dept, 1) <> "F" And (Text1.Text = "CFP" Or Text1.Text = "CPS" Or Text1.Text = "P" Or Text1.Text = "PS") Then
      m_CompNo = PUB_GetReceiptComp(pa(1), pa(2), pa(3), pa(4))
      '智權公司(patent信箱)
      If m_CompNo = "J" Then
         iPicNo = 21
         iPicNo2 = 26
      '非智權公司都用專利法律(patent信箱)
      Else
         iPicNo = 19
      End If

   'Added by Morgan 2015/8/5 +外專中文格式 --Kimi
   'Modified by Morgan 2015/9/1 P案還是照舊
   'ElseIf Left(m_Dept, 1) = "F" And (Text1.Text = "FCP" Or Text1.Text = "P") Then
   ElseIf Text1.Text = "FCP" Then
      iPicNo = 5
      iPicNo2 = 9

   End If
   'end 2013/12/27
      
End If 'Added by Morgan 2020/3/25

   
   bolRetry = False
    
On Error GoTo ERRORSECTION1
   
    If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
    If g_LetterDebug Then g_WordAp.Visible = True 'Added by Morgan 2019/8/12
    g_WordAp.Documents.add
    With g_WordAp
      'Add by Morgan 2013/1/8
      '切換為整頁模式,信頭才會正常顯示
      If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
         .ActiveWindow.ActivePane.View.Type = wdPageView
      Else
         .ActiveWindow.View.Type = wdPageView
      End If
      'end 2013/1/8
      
      .Selection.Font.Name = "標楷體"
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.Orientation = wdTextOrientationHorizontal
      .Selection.Font.Size = 14
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(2)
      'Modify by Morgan 2008/7/3
      '.Selection.PageSetup.TopMargin = .CentimetersToPoints(3.53)
      '.Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(4.1)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(2.5)
      'end 2008/7/3
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
      .Selection.ParagraphFormat.DisableLineHeightGrid = True
      'Add by Morgan 2008/7/17 配合新的開窗定稿改固定行高
      .Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
      .Selection.ParagraphFormat.LineSpacing = 15
      'end 2008/7/17
      
      'Add by Morgan 2010/12/13
      '信函信頭
      If txtLetterHead <> "N" Then
         'Added by Morgan 2013/12/27
         If iPicNo2 > 0 Then
            If PUB_ReadDB2File(stFileName, iPicNo) = True Then
               .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
               Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
               oShape.ZOrder 4
               oShape.LockAnchor = True
               oShape.LockAspectRatio = -1
               oShape.Width = .CentimetersToPoints(21)
               oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
               oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
               oShape.Left = .CentimetersToPoints(0)
               'Added by Morgan 2020/3/26 統一
               If strSrvDate(1) >= 智慧所更名日 Then
                  oShape.Top = .CentimetersToPoints(0)
               Else
               'end 2020/3/26
                  oShape.Top = .CentimetersToPoints(0.5)
               End If 'Added by Morgan 2020/3/26
               If PUB_ReadDB2File(stFileName, iPicNo2) = True Then
                  .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
                  Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
                  oShape.ZOrder 4
                  oShape.LockAnchor = True
                  oShape.LockAspectRatio = -1
                  oShape.Width = .CentimetersToPoints(21)
                  oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                  oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                  oShape.Left = .CentimetersToPoints(0)
                  'Added by Morgan 2020/3/26
                  If strSrvDate(1) >= 智慧所更名日 Then
                     oShape.Top = .CentimetersToPoints(27.6)
                  Else
                  'end 2020/3/26
                     oShape.Top = .CentimetersToPoints(27)
                  End If 'Added by Morgan 2020/3/26
               End If
               .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
               .Selection.EndKey Unit:=wdStory
            End If
         Else
         'end 2013/12/27
            If PUB_ReadDB2File(stFileName, iPicNo) = True Then
               '插入圖片檔案
               'Modified by Morgan 2011/12/13 改用物件變數控制因為Word2007預設的物件名稱不同會造成錯誤
               Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
               oShape.ZOrder 4
               oShape.LockAnchor = True
               oShape.LockAspectRatio = -1
               oShape.Width = 546.5
               oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
               oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
               oShape.Left = .CentimetersToPoints(1)
               'Modified by Morgan 2015/6/22 配合E化列印EMail信頭向上微調
               'oShape.Top = .CentimetersToPoints(1)
               'Added by Morgan 2020/3/26 統一
               If strSrvDate(1) >= 智慧所更名日 Then
                  oShape.Top = .CentimetersToPoints(0)
               Else
               'end 2020/3/26
                  oShape.Top = .CentimetersToPoints(0.8)
               End If 'Added by Morgan 2020/3/26
               oShape.WrapFormat.Type = wdWrapNone
               .Selection.EndKey Unit:=wdStory
            End If
         End If 'Added by Morgan 2013/12/27
      End If
      
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeParagraph
     
      If Option2.Value Then
         stReceiver = Trim(fa(0))
         stAddr = Trim(fa(16))
         '2008/10/30 add by sonia
         If Option4.Value Then
            stContact = Combo4
            If Combo5 <> "" Then stContact = stContact & vbCrLf & Combo5 'Added by Morgan 2017/1/19 +聯絡人2
         ElseIf Option5.Value Then
            stContact = Combo6
         End If
         '2008/10/30 end
      ElseIf Option6.Value Then
         stReceiver = Trim(cfa(0))
         stAddr = Trim(cfa(16))
         '2008/10/30 add by sonia
         If Option4.Value Then
            stContact = Combo4
            If Combo5 <> "" Then stContact = stContact & vbCrLf & Combo5 'Added by Morgan 2017/1/19 +聯絡人2
         ElseIf Option5.Value Then
            stContact = Combo6
         End If
         '2008/10/30 end
      ElseIf Option3.Value Then
         stReceiver = Trim(cu(0))
         stAddr = Trim(cu(16))
         stContact = Trim(m_Contact)
      End If
      m_MySt(1) = pa(1)
      m_MySt(2) = pa(2)
      m_MySt(3) = pa(3)
      m_MySt(4) = pa(4)
               
      'Add By Sindy 2009/10/15
      Select Case pa(1)
         Case "T", "TF", "CFT", "FCT"
               ' 由系統種類對照檔取得系統別(名稱)
               strExc(0) = "SELECT SK02 FROM SYSTEMKIND WHERE SK01='" & pa(1) & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  basLetter.m_SysKind = RsTemp.Fields(0).Value
               Else
                  MsgBox "無法由系統種類對照檔取得系統別 !", vbCritical
                  Exit Sub
               End If
               basLetter.PaSt = "PA01='" & m_MySt(1) & "' AND PA02='" & m_MySt(2) & "' AND PA03='" & _
                                          m_MySt(3) & "' AND PA04='" & m_MySt(4) & "'"
               basLetter.TmSt = "TM01='" & m_MySt(1) & "' AND TM02='" & m_MySt(2) & "' AND TM03='" & _
                                          m_MySt(3) & "' AND TM04='" & m_MySt(4) & "'"
               basLetter.LcSt = "LC01='" & m_MySt(1) & "' AND LC02='" & m_MySt(2) & "' AND LC03='" & _
                                          m_MySt(3) & "' AND LC04='" & m_MySt(4) & "'"
               basLetter.HcSt = "HC01='" & m_MySt(1) & "' AND HC02='" & m_MySt(2) & "' AND HC03='" & _
                                          m_MySt(3) & "' AND HC04='" & m_MySt(4) & "'"
               basLetter.SpSt = "SP01='" & m_MySt(1) & "' AND SP02='" & m_MySt(2) & "' AND SP03='" & _
                                          m_MySt(3) & "' AND SP04='" & m_MySt(4) & "'"
               'Modify By Sindy 2021/10/8 取消 中申開窗郵號/掛 判斷,因不知為什麼用"中申"
               '.Selection.TypeText IIf(Trim(ExceptFieldData2("中申開窗郵號/掛")) <> "", String(14, "　") & "掛號", "")
               .Selection.TypeText String(14, "　") & "掛號"
               '2021/10/8 END
               .Selection.TypeParagraph
               .Selection.TypeText ExceptFieldData2("中代開窗地址")
               .Selection.TypeParagraph
               .Selection.TypeParagraph
'               'Add by Morgan 2007/1/22 傳真號碼
'               strExc(0) = m_strFax(0)
'               If m_strFax(2) <> "" Then
'                  If m_strFax(0) <> "" Then
'                     strExc(0) = strExc(0) & ", " & m_strFax(2)
'                  Else
'                     strExc(0) = m_strFax(2)
'                  End If
'               End If
'               .Selection.TypeText "ＦＡＸ：" & strExc(0)
'               .Selection.TypeParagraph
               .Selection.TypeText "發函日期：" & ExceptFieldData("系統日/中西")
               .Selection.TypeParagraph
               '彼所案號
               Select Case pa(1) '判斷系統類別
                  Case "P", "CFP", "FCP" '專利
                     strText = pa(77)
                  Case "T", "TF", "CFT", "FCT" '商標
                     strText = pa(45)
                  'modify by sonia 2019/7/30 +ACS系統類別
                  Case "L", "FCL", "CFL", "LIN", "ACS" '法務
                     strText = pa(23)
                  Case "LA" '顧問
                     strText = ""
                  Case Else '服務業務
                     strText = pa(27)
               End Select
               .Selection.TypeText "貴方卷號：" & strText
               .Selection.TypeParagraph
               .Selection.TypeText "我方案號：" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & IIf(pa(4) = "00", "", "-" & pa(4)))
               .Selection.TypeParagraph
               .Selection.TypeParagraph
               If Left(Trim(m_Combo8), 1) = "2" Then
                  .Selection.TypeText "Re：通知台灣商標註冊核駁前先行通知"
               Else
                  .Selection.TypeText "Re：通知台灣商標"
               End If
               .Selection.TypeParagraph
               .Selection.TypeText "　　申請人：" & GetCustName(Text1.Text & Text2.Text & Text3.Text & Text4.Text, "1", True, "　　　　")
               .Selection.TypeParagraph
               .Selection.TypeText "　　商標：" & pa(5)
               .Selection.TypeParagraph
               .Selection.TypeText "　　類別：第" & pa(9) & "類"
               .Selection.TypeParagraph
               .Selection.TypeText "　　" & ExceptFieldData("(商標號數標題)") & "：" & ExceptFieldData("(商標號數)")
               .Selection.TypeParagraph
               .Selection.TypeParagraph
         '2009/10/15 End
         Case Else
               iLineCount = 0
               '地址
               If stAddr <> "" Then
                  strExc(0) = stAddr
                  Do While Len(strExc(0)) > 17
                     .Selection.TypeText Left(strExc(0), 17)
                     .Selection.TypeParagraph
                     strExc(0) = Mid(strExc(0), 18)
                     iLineCount = iLineCount + 1
                  Loop
                  If strExc(0) <> "" Then
                     .Selection.TypeText strExc(0)
                     .Selection.TypeParagraph
                     iLineCount = iLineCount + 1
                  End If
               End If
         
               '收件人
               If stReceiver <> "" Then
                  strExc(0) = stReceiver
                  Do While Len(strExc(0)) > 17
                     .Selection.TypeText Left(strExc(0), 17)
                     .Selection.TypeParagraph
                     strExc(0) = Mid(strExc(0), 18)
                     iLineCount = iLineCount + 1
                  Loop
                  If Len(strExc(0)) < 5 Then
                     .Selection.TypeText strExc(0) & "　　" & stContact & String(4 - Len(strExc(0)), "　") & "鈞啟(" & CaseNo & ")"
                     .Selection.TypeParagraph
                     iLineCount = iLineCount + 1
                  Else
                     .Selection.TypeText strExc(0)
                     .Selection.TypeParagraph
                     iLineCount = iLineCount + 1
                     .Selection.TypeText stContact & String(6, "　") & "鈞啟(" & CaseNo & ")"
                     .Selection.TypeParagraph
                     iLineCount = iLineCount + 1
                  End If
               Else
                  .Selection.TypeText stContact & String(6, "　") & "鈞啟(" & CaseNo & ")"
                  .Selection.TypeParagraph
                  iLineCount = iLineCount + 1
               End If
         
               'Added by Lydia 2020/08/26 (限外專)受文者非X編號，不論是FC或CF代理人以及語言都帶備註出來；其他部門的都先不動，將來有需要再加。
               If strSrvDate(1) >= 各項指示啟用日 And Left(m_StrUserST03, 2) = "F2" And (Option2.Value = True Or Option6.Value = True) And Text5 & Text6(0) & Text6(1) <> "" Then
                    .Selection.TypeParagraph
                    PrintMemo
                   .Selection.TypeParagraph
               Else
               'end 2020/08/26
                    '補滿5行
                    For intI = iLineCount + 1 To 5
                       .Selection.TypeParagraph
                    Next
               End If 'Added by Lydia 2020/08/26
               
               'Add by Morgan 2006/10/11列距上下加0.1cm --郭
               .Selection.ParagraphFormat.SpaceBefore = .CentimetersToPoints(0.1)
               .Selection.ParagraphFormat.SpaceAfter = .CentimetersToPoints(0.1)
               .Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
               'end 2006/10/11
         
               'Add by Morgan 2007/1/22 傳真號碼
               'Modified by Morgan 2016/8/5 P,FCP 案不印傳真 --郭雅娟,邱子瑜
               If pa(1) <> "FCP" And pa(1) <> "P" Then
                  strExc(0) = m_strFax(0)
                  If m_strFax(2) <> "" Then
                     If m_strFax(0) <> "" Then
                        strExc(0) = strExc(0) & ", " & m_strFax(2)
                     Else
                        strExc(0) = m_strFax(2)
                     End If
                  End If
                  .Selection.TypeParagraph
                  .Selection.TypeText "ＦＡＸ：" & strExc(0)
               End If
               'end 2007/1/22
               '帶出系統日期
               .Selection.TypeParagraph
               .Selection.TypeText "發函日期：" & Mid(strSrvDate(1), 1, 4) & "年" & Mid(strSrvDate(1), 5, 2) & "月" & Mid(strSrvDate(1), 7, 2) & "日"
               '彼所案號
               .Selection.TypeParagraph
               If Option2.Value = True Then
                  'Modified by Morgan 2015/8/5 +FCP
                  If pa(1) = "P" Or pa(1) = "CFP" Or pa(1) = "FCP" Then
                     .Selection.TypeText "貴方案號：" & pa(77)
                  '2008/10/29 ADD BY SONIA
                  ElseIf pa(1) = "T" Or pa(1) = "TF" Or pa(1) = "CFT" Then
                     .Selection.TypeText "貴方案號：" & pa(45)
                  ElseIf pa(1) = "L" Or pa(1) = "CFL" Then
                     .Selection.TypeText "貴方案號：" & pa(23)
                  '2008/10/29 END
                  Else
                     .Selection.TypeText "貴方案號：" & pa(27)
                  End If
               ElseIf Option6.Value = True Then
                  'Modify by Morgan 2011/1/5 +CP57 is null,-CP27>0
                  strExc(0) = "select CP45 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                                  " AND CP44 Is Not Null AND cp57 is null AND CP09<'C' Order By CP27 Desc, CP09 Desc "
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     .Selection.TypeText "貴方案號：" & RsTemp.Fields(0)
                  Else
                     .Selection.TypeText "貴方案號："
                  End If
               End If
               '本所案號
               .Selection.TypeParagraph
               .Selection.TypeText "我方案號：" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & IIf(pa(4) = "00", "", "-" & pa(4)))
               '案件名稱
               .Selection.TypeParagraph
               .Selection.TypeText "名稱：" & pa(5)
               '申請案號/審定號
               '2008/10/29 MODIFY BY SONIA
               '.Selection.TypeParagraph
               '.Selection.TypeText "申請號：" & pa(11)
               'Modified by Morgan 2015/8/5 +FCP
               If pa(1) = "P" Or pa(1) = "CFP" Or pa(1) = "FCP" Then
                  .Selection.TypeParagraph
                  .Selection.TypeText "申請號：" & pa(11)
               ElseIf pa(1) = "T" Or pa(1) = "TF" Or pa(1) = "CFT" Then
                  .Selection.TypeParagraph
                  If pa(15) = "" Then
                     .Selection.TypeText "申請號：" & pa(12)
                  Else
                     .Selection.TypeText "審定號：" & pa(15)
                  End If
               End If
               '2008/10/29 END
               '申請人
               .Selection.TypeParagraph
               .Selection.TypeText "申請人：" & GetCustName(Text1.Text & Text2.Text & Text3.Text & Text4.Text, "1", True, "　　　　")
               .Selection.TypeParagraph
               .Selection.TypeParagraph
      End Select
            
      .Selection.TypeText "敬啟者："
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      '2008/10/29 MODIFY BY SONIA
      '.Selection.TypeText "　　前述專利案，"
      'Modified by Morgan 2015/8/5 +FCP
      If pa(1) = "P" Or pa(1) = "CFP" Or pa(1) = "FCP" Then
         .Selection.TypeText "　　前述專利案，"
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.TypeParagraph
      ElseIf pa(1) = "T" Or pa(1) = "TF" Or pa(1) = "CFT" Or pa(1) = "FCT" Then
         
         'Add By Sindy 2009/10/15
         Select Case Left(m_Combo8, 1)
            Case "2"             '核駁前先行通知1202
               StrSQLa = "Select CP07,CP08,CP36,NVL(NVL(CP37,CP38),CP39),NVL(NVL(TG06||TG15,TG07||TG16),TG08||TG17),CP06 FROM CaseProgress,TMGOODS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                         " AND CP10='1202' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' AND CP01=TG01(+) AND CP02=TG02(+) AND CP03=TG03(+) AND CP04=TG04(+) "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  m_CP07 = Mid(rsA.Fields(0).Value, 1, 4) & "年" & Mid(rsA.Fields(0).Value, 5, 2) & "月" & Mid(rsA.Fields(0).Value, 7, 2) & "日"
'                  m_DueDT = ChangeWDateStringToWString(DateAdd("D", -7, ChangeWStringToWDateString(rsA.Fields("CP06").Value)))
'                  m_CP06 = Mid(m_DueDT, 1, 4) & "年" & Mid(m_DueDT, 5, 2) & "月" & Mid(m_DueDT, 7, 2) & "日"
                  'Modify By Sindy 2012/4/23
                  m_CP06 = Mid(rsA.Fields("CP06").Value, 1, 4) & "年" & Mid(rsA.Fields("CP06").Value, 5, 2) & "月" & Mid(rsA.Fields("CP06").Value, 7, 2) & "日"
                  '本所期限-7日
                  m_DueDT = CompDate(2, -7, rsA.Fields("CP06").Value)
                  m_DueDate = Mid(m_DueDT, 1, 4) & "年" & Mid(m_DueDT, 5, 2) & "月" & Mid(m_DueDT, 7, 2) & "日"
                  '2012/4/23 End
                  'add by sonia 2021/4/20 多商品類別的商品也要都抓T-231208,下面定稿內容程式rsA.Fields(4)換成m_TMGoods
                  m_TMGoods = ""
                  Do While Not rsA.EOF
                     If m_TMGoods <> "" Then m_TMGoods = m_TMGoods & "；"
                     m_TMGoods = m_TMGoods & rsA.Fields(4)
                     rsA.MoveNext
                  Loop
                  rsA.MoveFirst
                  'end 2021/4/20
               End If
               If pa(10) = "000" Then    '大->台
                  Select Case m_Combo8
                     Case "21"      'T核駁前先行通知1202第23條第1項第1款(AXX)及第23條第1項第2款(BXX)
                        .Selection.TypeText "　　有關上述商標註冊申請案，頃接智慧財產局之核駁理由先行通知書，謂本件商標圖樣上之「" & CASENAME & "」部分，為指定商品之相關說明，不具識別性，有違商標法第5條第2項及第23條第1項第1款、第2款之規定，隨函檢附" & rsA.Fields("CP08").Value & "核駁理由先行通知書，請查照。"
                        .Selection.TypeParagraph
                        .Selection.TypeText "　　智慧局認為，本件商標圖樣上之「" & CASENAME & "」為　　　之意，以之作為商標，指定使用於「" & m_TMGoods & "」等商品，為指定商品之相關說明，難謂足以表彰指定之商品並與他人商品相區別，不具識別性。如申請人欲答辯爭取，須檢附長時間大量使用該商標之廣告、行銷及各國的註冊資料等，證明足資作為消費者辨識商品來源之依據，已具識別性，以爭取之。如申請人欲答辯爭取，費用為人民幣1900元。如欲聲明圖樣中「" & CASENAME & "」部分不在專用之列，費用為人民幣500元。"
                     Case "22"      'T核駁前先行通知1202第23條第1項第2款(BXX)及第23條第1項第11款(FXX)
                        .Selection.TypeText "　　有關上述商標註冊申請案，頃接智慧財產局之核駁理由先行通知書，謂本件商標圖樣上之「" & CASENAME & "」，有違商標法第23條第1項第2款及第11款之規定，隨函檢附" & rsA.Fields("CP08").Value & "核駁理由先行通知書，請查照。"
                        .Selection.TypeParagraph
                        .Selection.TypeText "　　智慧局認為，以「" & CASENAME & "」作為商標之一部分，指定使用於「" & m_TMGoods & "」等商品，直接明顯為指定商品之相關說明；指定使用於「" & m_TMGoods & "」等商品，則有使相關消費大眾對商品之性質產生誤信誤認。如申請人欲答辯爭取時，必須檢附長時間大量使用該商標之廣告、行銷資料，證明無使相關消費大眾對商品之性質產生誤信誤認，答辯費用為人民幣1900元或聲明圖樣中「" & CASENAME & "」部分不在專用之列，並刪除「" & m_TMGoods & "」等商品，費用為人民幣500元。"
                     Case "23"      'T核駁前先行通知1202第23條第1項第2款(BXX)及第23條第1項第13款(HXX)
                        .Selection.TypeText "　　有關上述商標註冊申請案，頃接智慧財產局之核駁理由先行通知書，謂本件商標與註冊第" & rsA.Fields(2) & "號、註冊第" & rsA.Fields(2) & "號「" & rsA.Fields(3) & "」商標中文近似；而商標圖樣上之「" & CASENAME & "」為指定商品之說明，有違商標法第23條第1項第2款及第13款之規定，隨函檢附" & rsA.Fields("CP08") & "核駁理由先行通知書及引証商標圖樣，請查照。"
                        .Selection.TypeParagraph
                        .Selection.TypeText "　　於申請前，曾告知有近似之商標註冊在先，今智慧局亦認為與該商標構成近似，由於本件商標圖樣上之「" & CASENAME & "」，與前揭註冊商標「" & CASENAME & "」，且指定使用於相同或類似之商品，依現行商標審查基準，係屬近似之商標，爭取應有困難，申請人可刪除與前述商標類似之「" & m_TMGoods & "」等商品，以克服核駁理由。如申請人欲答辯爭取，費用為人民幣1900元；如欲刪除商品，費用為人民幣500元。"
                        .Selection.TypeParagraph
                        .Selection.TypeText "　　另，圖樣上之「" & CASENAME & "」部分，為指定商品之相關說明，不具識別性，應聲明圖樣中「" & CASENAME & "」部分不在專用之列。"
                     Case "24"      'T核駁前先行通知1202第23條第1項第11款(FXX)及第23條第1項第13款(HXX)
                        .Selection.TypeText "　　有關上述商標註冊申請案，頃接智慧財產局之核駁理由先行通知書，謂本件商標圖樣上之「" & CASENAME & "」，與註冊第" & rsA.Fields(2) & "號「" & rsA.Fields(3) & "」商標近似，有違商標法第23條第1項第13款之規定，而圖樣中「" & CASENAME & "」，則有違商標法第23條第1項第11款之規定，隨函檢附" & rsA.Fields("CP08") & "核駁先行通知書及引証商標圖樣，請查照。"
                        .Selection.TypeParagraph
                        .Selection.TypeText "　　由於本件商標與前述商標構成近似，且指定使用於同一及類似商品，依現行商標審查基準，應屬近似之商標。而圖樣中「" & CASENAME & "」，則有使一般消費大眾對指定商品之產地、來源產生誤信誤認之虞。如申請人欲答辯爭取時，費用為美金人民幣1900元。"
                     Case "25"      'T核駁前先行通知1202第23條第1項第1款(AXX)
                        .Selection.TypeText "　　有關上述商標註冊申請案，頃接智慧財產局之核駁理由先行通知書，謂本件商標圖樣上之「" & CASENAME & "」為指定服務之說明，不具識別性，有違商標法第5條第2項及第23條第1項第1款之規定，隨函檢附" & rsA.Fields("CP08").Value & "核駁理由先行通知書，請查照。"
                        .Selection.TypeParagraph
                        .Selection.TypeText "　　智慧局認為，本件商標圖樣上之「" & CASENAME & "」，為　　　之意，以之作為商標，使用於指定之服務，為指定服務之相關說明，不具識別性。若申請人欲答辯爭取，須檢附長時間大量使用該商標之廣告、行銷資料及各國的商標註冊等資料，證明足資作為消費者辨識服務來源之依據，已具識別性，以爭取之。如申請人欲答辯爭取，費用為人民幣1900元。"
                     Case "26"      'T核駁前先行通知1202'第23條第1項第2款(BXX)
                        .Selection.TypeText "　　有關上述商標註冊申請案，頃接智慧財產局之核駁理由先行通知書，謂本件商標圖樣上之「" & CASENAME & "」為指定服務之說明，有違商標法第23條第1項第2款之規定，隨函檢附" & rsA.Fields("CP08").Value & "核駁理由先行通知書，請查照。"
                        .Selection.TypeParagraph
                        .Selection.TypeText "　　智慧局認為，本件商標圖樣上之「" & CASENAME & "」為　　　，以之作為商標之一部分，使用於指定之服務，為指定服務之相關說明，不具識別性。若申請人欲答辯爭取，須檢附長時間大量使用該商標之廣告、行銷資料及各國的商標註冊等資料，證明足資作為消費者辨識服務來源之依據，已具識別性；或聲明圖樣中「" & CASENAME & "」部分不在專用之列，以爭取之。如申請人欲答辯爭取，費用為人民幣1900元；如欲聲明圖樣中「" & CASENAME & "」部分不在專用之列，費用為人民幣500元。"
                     Case "27"      'T核駁前先行通知1202第23條第1項第11款(FXX)
                        .Selection.TypeText "　　有關上述商標註冊申請案，頃接智慧財產局之核駁理由先行通知書，謂本件商標圖樣上之「" & CASENAME & "」，有違商標法第23條第1項第11款之規定，隨函檢附" & rsA.Fields("CP08").Value & "核駁理由先行通知書，請查照。"
                        .Selection.TypeParagraph
                        .Selection.TypeText "　　智慧局認為，以「" & CASENAME & "」作為商標之一部分，指定使用於「" & m_TMGoods & "」等商品，有使相關消費大眾對其表彰商品之性質產生誤信誤認。如申請人欲答辯爭取時，必須檢附長時間大量使用該商標之廣告、行銷資料，證明無使相關消費大眾對商品之性質產生誤信誤認，答辯費用為人民幣1900元。或刪除圖樣中「" & CASENAME & "」部分，費用為人民幣500元。"
                     Case "29"      'T核駁前先行通知1202第23條第1項第13款(HXX)
                        .Selection.TypeText "　　有關上述商標註冊申請案，頃接智慧財產局之核駁理由先行通知書，謂本件商標與註冊第" & rsA.Fields(2).Value & "號「" & rsA.Fields(3).Value & "」商標中文相同，有違商標法第23條第1項第13款之規定，隨函檢附" & rsA.Fields("CP08").Value & "核駁理由先行通知書及引証商標圖樣，請查照。"
                        .Selection.TypeParagraph
                        .Selection.TypeText "　　於申請前，曾告知有近似之商標註冊在先，今智慧局亦認為與該商標構成近似，由於本件商標與前述商標皆有相同之中文「" & CASENAME & "」，且指定使用於同一及類似商品，依現行商標審查基準，係屬近似之商標，爭取應有困難，申請人可刪除與前述商標類似之「" & m_TMGoods & "」商品，保留「」等商品，以克服核駁理由。如申請人欲答辯爭取，費用為人民幣1900元；如欲刪除商品，費用為人民幣500元。"
                     Case "2B"      'T核駁前先行通知1202第23條第1項第15款(JXX)
                        .Selection.TypeText "　　有關上述商標註冊申請案，頃接智慧財產局之核駁理由先行通知書，謂本件商標圖樣上之「人像圖」係為他人之肖像，有違商標法第23條第1項第15款之規定，隨函檢附" & rsA.Fields("CP08").Value & "核駁理由先行通知書，請查照。"
                        .Selection.TypeParagraph
                        .Selection.TypeText "　　由於本件商標商標圖樣上之「人像圖」係為他人之肖像，依台灣現行商標法規定，須徵得其同意，方可取得商標註冊。經本所與審查員溝通，告知該商標圖樣已於大陸取得商標註冊，應可證明肖像所有人已同意申請人以其肖像申請商標註冊，惟審查員謂此只能證明肖像所有人同意其於大陸申請商標註冊，不能證明其亦同意在台灣申請商標註冊，故請檢附肖像所有人的同意書。"
                     Case "2C"      'T核駁前先行通知1202第23條第1項第1款(AXX)及第23條第1項第2款(BXX)及第23條第1項第11款(FXX)
                        .Selection.TypeText "　　有關上述商標註冊申請案，頃接智慧財產局之核駁理由先行通知書，謂本件商標圖樣上之「" & CASENAME & "」為指定商品之說明，有違商標法第5條第2項及第23條第1項第1、2、11款之規定，隨函檢附" & rsA.Fields("CP08").Value & "核駁理由先行通知書，請查照。"
                        .Selection.TypeParagraph
                        .Selection.TypeText "　　智慧局認為，以「" & CASENAME & "」作為商標之一部分，使用於指定之商品，為指定商品之相關說明，不具識別性，且有致一般消費大眾對其所表彰商品之產地、來源產生誤信誤認。如申請人欲答辯爭取，須檢附長時間大量使用該商標之廣告、行銷及各國的註冊資料等，證明足資作為消費者辨識商品來源之依據，已具識別性，且無使相關消費大眾對商品之產地產生誤信誤認；或聲明圖樣中「" & CASENAME & "」部分不在專用之列、並刪除圖樣中之「" & CASENAME & "」部分，以爭取之。如申請人欲答辯爭取，費用為人民幣1900元；如欲聲明圖樣中「" & CASENAME & "」部分不在專用之列、並刪除圖樣中之「" & CASENAME & "」部分，費用為人民幣500元。"
                  End Select
                  If rsA.State <> adStateClosed Then rsA.Close
                  Set rsA = Nothing
                  .Selection.TypeParagraph
                  'Modify By Sindy 2012/4/23
                  '.Selection.TypeText "　　依法申請人須" & m_CP07 & "以前回覆智慧財產局，若未於期限內回覆，本件商標將遭駁回處分，煩請於" & m_CP06 & "前告知續行方式，以利本所作業。"
                  .Selection.TypeText "　　依法申請人須" & m_CP06 & "以前回覆智慧財產局，若未於期限內回覆，本件商標將遭駁回處分，煩請於" & m_DueDate & "前告知續行方式，以利本所作業。"
                  .Selection.TypeParagraph
                  .Selection.TypeText "　　若尚有任何問題，請隨時與本所聯繫。"
                  .Selection.TypeParagraph
               End If
            '2009/10/15 End
            'ADD BY SONIA 2015/9/16 台灣案提申請意見書後之申請或分割核駁
            Case "5"   '1002
               If m_Combo8 = "56" Then
                  m_CP49 = ""
                  StrSQLa = "Select CP07,CP06,CP49 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " AND CP10='1002' AND CP27 IS NULL AND CP57 IS NULL AND CP09>'C' "
                  rsA.CursorLocation = adUseClient
                  rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                  If rsA.RecordCount > 0 Then
                     m_CP07 = Mid(rsA.Fields(0).Value, 1, 4) & "年" & Mid(rsA.Fields(0).Value, 5, 2) & "月" & Mid(rsA.Fields(0).Value, 7, 2) & "日"
                     m_CP06 = Mid(rsA.Fields("CP06").Value, 1, 4) & "年" & Mid(rsA.Fields("CP06").Value, 5, 2) & "月" & Mid(rsA.Fields("CP06").Value, 7, 2) & "日"
                     '本所期限-3工作天
                     m_DueDT = CompWorkDay(3, rsA.Fields("CP06").Value, 1)
                     m_DueDate = Mid(m_DueDT, 1, 4) & "年" & Mid(m_DueDT, 5, 2) & "月" & Mid(m_DueDT, 7, 2) & "日"
                     m_CP49 = "" & rsA.Fields("CP49").Value
                     GetLaw   '依條款代碼取得條款名稱caselaw
                  End If
                  g_WordAp.Selection.TypeText "　　有關上述商標註冊申請案經提出意見書後，頃接經濟部智慧財產局" & AppNo & "審定書，謂本案與商標法規定不合，予以核駁，茲隨函檢附核駁審定正本乙份，敬請查收。"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　本案前經智慧財產局認有違反商標法" & caselaw & "規定之嫌，雖經我方說明(提出)            但未獲採納。本案續行爭取的方向可"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　本案係有關商標法" & caselaw & "規定之爭執，我方前於意見書中主要強調　　　惟經智慧財產局認為　　　　。申請人如有意續爭，可"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　商標法" & caselaw & "規定之適用係以　為要件之一，故我方前於意見書中，係主張　　　　　但智慧財產局仍認　　　。針對核駁理由，申請人宜"
                  g_WordAp.Selection.TypeParagraph
                  g_WordAp.Selection.TypeText "　　本案依法應於經濟部智慧財產局文到次日起 30 日內提出訴願（即" & m_CP06 & "以前）否則視為結案。如申請人對上項核駁處分如有不服欲提出訴願，請於" & m_DueDate & "以前與本所聯繫，以共商提出事宜。"
                  g_WordAp.Selection.TypeParagraph
                  If rsA.State <> adStateClosed Then rsA.Close
                  Set rsA = Nothing
               End If
            'END 2015/9/16
            Case Else
               .Selection.TypeText "　　前述商標案，"
               .Selection.TypeParagraph
               .Selection.TypeParagraph
               .Selection.TypeParagraph
         End Select
         
      ElseIf pa(1) = "L" Or pa(1) = "CFL" Then
         .Selection.TypeText "　　前述法務案，"
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.TypeParagraph
      Else
         .Selection.TypeText "　　前述案件，"
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.TypeParagraph
      End If
      '2008/10/29 END
      'Remove by Morgan 2006/12/13 以免會跳頁且不用手動刪除--玲玲
      '.Selection.TypeParagraph
      '.Selection.TypeParagraph
      '.Selection.TypeParagraph
      '.Selection.TypeParagraph
      '.Selection.TypeParagraph
      '.Selection.TypeParagraph
      'end 2006/12/13
      'Remove by Morgan 2006/10/11 因為列距有加高
      '.Selection.TypeParagraph
      '.Selection.TypeParagraph
      '.Selection.TypeParagraph
      '.Selection.TypeParagraph
      '.Selection.TypeParagraph
      '.Selection.TypeParagraph
      'end 2006/10/11
      .Selection.TypeText "　　耑此　　順頌"
      .Selection.TypeParagraph
      'Modify By Sindy 2012/6/6
      '.Selection.TypeText "業　　祺"
      .Selection.TypeText "時祺"
      '2012/6/6 End
      .Selection.TypeParagraph
      'Added by Morgan 2014/11/13
      '簽名檔
      BolHaveImg = False
      'Added by Lydia 2017/05/03 P大陸案(大-台,台-大)改成公司別在上,簽名檔在下
      bolP020 = False
      'Modify by Amy 2018/07/27  所有取得信頭GetImgByteFile 改成用 PUB_ReadDB2File
      If pa(1) = "P" And (Option6.Value = True Or Option2.Value = True) Then
           bolP020 = True
           .Selection.TypeParagraph
           .Selection.TypeText PUB_GetCompany(pa(1), pa(2), pa(3), pa(4)) '公司別 敬上
           .Selection.TypeParagraph
           
           'Modified by Lydia 2022/03/29
           'If InStr(Replace(PUB_GetCompany(pa(1), pa(2), pa(3), pa(4)), "　", ""), "閻啟泰") > 0 Then
           StrSQLa = PUB_UniToBIG5(Replace(PUB_GetCompany(pa(1), pa(2), pa(3), pa(4)), "　", ""))
           If InStr(StrSQLa, "閻啟泰") > 0 Or InStr(StrSQLa, "閻?泰") > 0 Then
           'end 2022/03/29
               strExc(2) = "000034"
               strSignFile = App.path & "\" & "$$M51000034000.jpg"
               iPicPos = 雙署名POS_1
            'Modified by Morgan 2023/5/3
            'Else
            ElseIf InStr(StrSQLa, "郭雅娟") > 0 Then
               strExc(2) = "000092"
               strSignFile = App.path & "\" & "$$M51000092000.jpg"
               iPicPos = 雙署名POS_2
            ElseIf InStr(StrSQLa, "王錦寬") > 0 Then
            'end 2023/5/3
            
               strExc(2) = "000054"
               strSignFile = App.path & "\" & "$$M51000054000.jpg"
               iPicPos = 雙署名POS_2
           End If
           
           If strSignFile <> "" Then
               If Dir(strSignFile) = "" Then
                  BolHaveImg = PUB_ReadDB2File(strSignFile, Val(strExc(2)))
               Else
                  BolHaveImg = True
               End If
           End If
      Else
      'end 2017/05/03
        StrSQLa = PUB_UniToBIG5(Replace(PUB_GetCompany(pa(1), pa(2), pa(3), pa(4)), "　", "")) 'Added by Morgan 2023/5/3
        'Memo by Morgan 2023/5/3 P案已改雙署名，不會再執行到此段程式碼，無需修改
        If InStr(Replace(PUB_GetCompany(pa(1), pa(2), pa(3), pa(4)), "　", ""), "王錦寬") > 0 Then
           strSignFile = App.path & "\" & "$$M51000033000.jpg"
           If Dir(strSignFile) = "" Then
              BolHaveImg = PUB_ReadDB2File(strSignFile, 33)
           Else
              BolHaveImg = True
           End If
        'Added by Morgan 2014/12/12
        'Modified by Lydia 2022/03/29
        'ElseIf InStr(Replace(PUB_GetCompany(pa(1), pa(2), pa(3), pa(4)), "　", ""), "閻啟泰") > 0 Then
        'StrSQLa = PUB_UniToBIG5(Replace(PUB_GetCompany(pa(1), pa(2), pa(3), pa(4)), "　", "")) 'Removed by Morgan 2023/5/3 移到外層才有用
        ElseIf InStr(StrSQLa, "閻啟泰") > 0 Or InStr(StrSQLa, "閻?泰") > 0 Then
        'end 2022/03/29
           strSignFile = App.path & "\" & "$$M51000036000.jpg"
           If Dir(strSignFile) = "" Then
              BolHaveImg = PUB_ReadDB2File(strSignFile, 36)
           Else
              BolHaveImg = True
           End If
        End If
      End If 'end 2017/05/03
      'end 2018/07/26
      'end 2014/12/12
        
         If BolHaveImg Then
            'Added by Lydia 2017/05/03 P大陸案的簽名檔
            'Remove by Lydia 2017/05/22 簽名檔改塞空白
            'If bolP020 = True Then
            '    'Added by Lydia 2017/05/18 Word97若不顯示,無法取得正確位置,會停在信頭下方
            '    bVisibled = g_WordAp.Visible
            '    If bVisibled = False Then g_WordAp.Visible = True
            '    'end 2017/05/18
            '       Set oShape = .Selection.InlineShapes.AddPicture(FileName:=strSignFile, LinkToFile:=False, SaveWithDocument:=True).ConvertToShape
            '       pageY = .Selection.Information(wdVerticalPositionRelativeToPage) '記錄頁面位置
            '    If bVisibled = False Then g_WordAp.Visible = False 'Added by Lydia 2017/05/18 還原不顯示
            '
            '    oShape.LockAspectRatio = -1
            '    oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
            '    oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
            '    oShape.Top = pageY
            '    oShape.Left = .CentimetersToPoints(iPicPos)
            'Else
            'end 2017/05/03
                'Added by Lydia 2017/05/22 P大陸案的簽名檔位置，改塞空白的方式
                If bolP020 = True Then
                   '.Selection.TypeParagraph 'Remove by Lydia 2017/05/24
                   .Selection.TypeText String(iPicPos, "　")
                Else
                'end 2017/05/22
                  .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　　　　"
                End If 'end 2017/05/22
                
                intI = .Selection.ParagraphFormat.LineSpacingRule
                'Modified by Lydia 2017/05/24 排除P大陸指示信
                'If Val(.Version) >= 12 Then
                If Val(.Version) >= 12 And bolP020 = False Then
                   .Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
                   Set oShape = .Selection.InlineShapes.AddPicture(FileName:=strSignFile, LinkToFile:=False, SaveWithDocument:=True).ConvertToShape
                   oShape.WrapFormat.Type = 5
                   oShape.WrapFormat.Side = wdWrapBoth
                Else
                   '單行行高
                   .Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
                   .Selection.InlineShapes.AddPicture FileName:=strSignFile, LinkToFile:=False, SaveWithDocument:=True
                End If
            'End If 'end 2017/05/03
            
            .Selection.TypeParagraph
            .Selection.ParagraphFormat.LineSpacingRule = intI
         End If
      'End If 'Removed by Morgan 2014/12/12
      'strSignFile
      'end 2014/11/13
      .Selection.TypeParagraph
      '.Selection.TypeParagraph
      '.Selection.TypeParagraph
      '2008/10/29 modify by sonia
      '.Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　　王　錦　寬　敬上"
      '雙署名
      
      'Added by Morgan 2015/8/5 +外專中文格式 --Kimi
      '.Selection.TypeText PUB_GetCompany(pa(1), pa(2), pa(3), pa(4))
      If Text1.Text = "FCP" Then
         'Modified by Morgan 2020/3/25
         '.Selection.TypeText "　　　　　　　　　　　　　　　　　　　台一國際專利法律事務所"
         .Selection.TypeText "　　　　　　　　　　　　　　　　　　　" & CompNameQuery(m_CompNo)
         'end 2020/3/25
         .Selection.TypeParagraph
         'Modified by Lydia 2022/03/29 抓員工名稱
         '.Selection.TypeText "　　　　　　　　　　　　　　　　　　　副所長專利師　閻　啟泰"
         StrSQLa = GetStaffName("81040", True)
         'Modify By Sindy 2023/10/3 拿掉"副"字
         '.Selection.TypeText "　　　　　　　　　　　　　　　　　　　副所長專利師　" & Mid(StrSQLa, 1, 1) & "　" & Mid(StrSQLa, 2)
         .Selection.TypeText "　　　　　　　　　　　　　　　　　　　所長專利師　" & Mid(StrSQLa, 1, 1) & "　" & Mid(StrSQLa, 2)
         'end 2022/03/29
         .Selection.TypeParagraph
         .Selection.TypeText "　　　　　　　　　　　　　　　　　　　國外部專利處　顏　裕洋　敬上"
      'Modified by Lydia 2017/05/03 P大陸案改成公司別在上,簽名檔在下
      'Else
      ElseIf bolP020 = False Then
         .Selection.TypeText PUB_GetCompany(pa(1), pa(2), pa(3), pa(4))
      End If
      'end 2015/8/5
      
      '2008/10/29 end
      'Add by Amy 2013/08/23 大陸指示信+附註文字
      If Text1 = "P" Then
         bolBCP = ChkBCP() 'Added by Morgan 2023/9/12
         'Modified by Lydia 2017/05/03
         '.Selection.TypeParagraph
         If bolP020 = False Then .Selection.TypeParagraph
         'Modify By Sindy 2021/9/9
         '.Selection.TypeText "以下附註內容請承辦工程師依據個案案情予以保留或刪除。"
         '.Selection.TypeParagraph
         '.Selection.TypeText "附註：請 貴方依照我方所提供的意見(或修改本)準備正式陳述書，若陳述內容有實質性的補充或修改需交由我方確認，請貴方在補充或修正處以貫穿線表示刪除或以劃底線表示補充，並請於回函上說明。"
         'Modified by Morgan 2023/9/12
         '.Selection.TypeText "一、以下附註內容請承辦工程師依據個案案情予以保留或刪除。"
         .Selection.TypeText IIf(bolBCP, "一、", "") & "以下附註內容請承辦工程師依據個案案情予以保留或刪除。"
         'end 2023/9/12
         .Selection.TypeParagraph
         .Selection.TypeText "附註：請 貴方依照我方所提供的意見(或修改本)準備正式陳述書，若陳述內容有實質性的補充或修改需交由我方確認，請貴方在補充或修正處以貫穿線表示刪除或以劃底線表示補充，並請於回函上說明。"
                  
         If bolBCP Then 'Added by Morgan 2023/9/12
            .Selection.TypeParagraph
            .Selection.TypeParagraph
            .Selection.TypeText "二、若此程序為B類內部收文即表示為本所疏失或代理人疏失不會向客戶收取費用，則請依案情在指示信中帶入以下的內容："
            .Selection.TypeParagraph
            .Selection.TypeText "本所疏失－"
            .Selection.TypeParagraph
            .Selection.TypeText "由於本次補正(或修改)為本所作業疏失，無法向客戶收取費用，故請  貴方在代理費上給予本所最大優惠，謝謝。"
            .Selection.TypeParagraph
            .Selection.TypeText "代理人疏失－"
            .Selection.TypeParagraph
            .Selection.TypeText "由於本次補正(或修改)為  貴方處理上的疏失，本所無法向客戶收取費用，故請  貴方能自行吸收相關費用，謝謝。"
         End If 'Added by Morgan 2023/9/12
         
         '2021/9/9 END
      End If
      'end 2013/08/23
      'Add by Morgan 2007/2/7
      'Modified by Lydia 2019/09/27 改成共用
      'strExc(0) = GetLetterMemo(pa(1), "1")
      strExc(0) = Pub_GetLetterMemo(pa(1), "1")
      If strExc(0) <> "" Then
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.TypeText strExc(0)
      End If
      '2008/7/23 add by sonia 英數字改字體
      If Mid(m_Dept, 1, 2) = "P1" Then
         .Selection.WholeStory
         'modify by sonia 2019/11/22 郭同意改Arial為Times New Roman(與定稿統一),否則發文字號位置會跑掉
         .Selection.Font.Name = "Times New Roman"
      End If
      '2008/7/23 end
       
   End With
   
  g_WordAp.Visible = True
  g_WordAp.WindowState = wdWindowStateMaximize
  Set g_WordAp = Nothing 'Added by Morgan 2015/9/7

ERRORSECTION1:
   If Err.NUMBER <> 0 Then
      Select Case Err.NUMBER
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

'英文信
Private Sub WordEnglish()
   Dim stLetter As String
   Dim m As Integer, j As Integer
   Dim strCustName As String
   Dim strReceiver As String
   Dim bAllApp As Boolean
   Dim stFileName As String '暫存圖檔檔名
   Dim iPicNo As Integer, iPicNo2 As Integer '上,下圖檔代碼
   Dim oShape
   Dim bolShow As Boolean 'Added by Morgan 2024/3/29
   
   'Add by Morgan 2006/8/23
   '代理人名成兩欄並一行印
   If fa(1) <> "" And Len(fa(1)) <= 30 Then
      If fa(2) <> "" Then
         fa(1) = Trim(fa(1)) & " " & Trim(fa(2))
         fa(2) = ""
      End If
      If fa(3) <> "" Then
         fa(2) = Trim(fa(3)) & " " & Trim(fa(4))
         fa(3) = ""
         fa(4) = ""
      End If
   End If
   If cfa(1) <> "" And Len(cfa(1)) <= 30 Then
      If cfa(2) <> "" Then
         cfa(1) = Trim(cfa(1)) & " " & Trim(cfa(2))
         cfa(2) = ""
      End If
      If cfa(3) <> "" Then
         cfa(2) = Trim(cfa(3)) & " " & Trim(cfa(4))
         cfa(3) = ""
         cfa(4) = ""
      End If
   End If
   If cu(1) <> "" And Len(cu(1)) <= 30 Then
      If cu(2) <> "" Then
         cu(1) = Trim(cu(1)) & " " & Trim(cu(2))
         cu(2) = ""
      End If
      If cu(3) <> "" Then
         cu(2) = Trim(cu(3)) & " " & Trim(cu(4))
         cu(3) = ""
         cu(4) = ""
      End If
   End If
   
'Added by Morgan 2020/3/25
If strSrvDate(1) >= 智慧所更名日 Then
   m_CompNo = PUB_GetReceiptComp(pa(1), pa(2), pa(3), pa(4), True)
   PUB_GetLetterPicID m_CompNo, pa(1), iPicNo, iPicNo2, 2, True, m_Dept
Else
'end 2020/3/25

   
'Modify by Morgan 2011/7/6 改用新的信頭
'   Select Case Text1.Text
'      Case "FCT", "CFT"
'         iPicNo = 1
'      'Modify By Sindy 2009/07/24 增加LIN系統類別
'      Case "FCP", "FG", "FCL", "CFL", "P", "LIN"
'         iPicNo = 2
'      Case "CFP", "CPS"
'         iPicNo = 3
'      Case Else
'         iPicNo = 4
'   End Select
'   'end 2006/8/23
   iPicNo = 5
   iPicNo2 = 9
'end 2011/7/6
   
   'Added by Morgan 2013/12/27 改都要判斷,且改回傳公司別
   m_CompNo = PUB_GetReceiptComp(pa(1), pa(2), pa(3), pa(4))
   '智權公司
   If m_CompNo = "J" Then
      iPicNo = 21
      'patent 信箱
      If Left(m_Dept, 1) <> "F" And (Text1.Text = "CFP" Or Text1.Text = "CPS" Or Text1.Text = "P" Or Text1.Text = "PS") Then
         iPicNo2 = 22
      'ipdept 信箱
      Else
         iPicNo2 = 24
      End If
   '專利處對國外都用專利法律公司(patent信箱)
   ElseIf Left(m_Dept, 1) <> "F" And (Text1.Text = "CFP" Or Text1.Text = "CPS" Or Text1.Text = "P" Or Text1.Text = "PS") Then
      iPicNo2 = 18
   End If
   'end 2013/12/27
   
End If 'Added by Morgan 2020/3/25
   
   '是否印傳真封面
   'Modified by Lydia 2020/09/11 勾選各項指示
   If txtFaxFace <> "N" Then
      intI = vbYes
      If m_bEMail = True Then
         intI = MsgBox("本案有設定以EMail通知，是否要印傳真封面?", vbYesNo + vbDefaultButton2)
      End If
      If intI = vbYes Then
         Select Case Text1.Text
            Case "FCP"
               NowPrint Text1.Text & Text2.Text & Text3.Text & Text4.Text & "&000", "04", "99", False, strUserNum, , , True, stLetter, 1, Label11.Caption
            Case Else
               NowPrint Text1.Text & Text2.Text & Text3.Text & Text4.Text & "&000", "01", "99", False, strUserNum, , , True, stLetter, 1, Label11.Caption
         End Select
         If stLetter <> "" Then stLetter = stLetter & Chr(12)
         '因撰寫信函傳本所號&000,所以進度檔相關欄位不會有資料;TO:,FAX:,TEL:後面一定是空的
         'Modify by Morgan 2006/3/24 發信對象與信函內一致
         If Option2.Value = True And fa(1) <> "" Then
            strReceiver = fa(1)
            For j = 2 To 4
               If fa(j) <> "" Then
                  strReceiver = strReceiver & vbCrLf & String(4, " ") & fa(j)
               End If
            Next j
            stLetter = Replace(stLetter, "TO: ", "TO: " & strReceiver, 1, 1)
         ElseIf Option6.Value = True And cfa(1) <> "" Then
            strReceiver = cfa(1)
            For j = 2 To 4
               If cfa(j) <> "" Then
                  strReceiver = strReceiver & vbCrLf & String(4, " ") & cfa(j)
               End If
            Next j
            stLetter = Replace(stLetter, "TO: ", "TO: " & strReceiver, 1, 1)
         ElseIf Option3.Value = True And cu(1) <> "" Then
            strReceiver = cu(1)
            For j = 2 To 4
               If cu(j) <> "" Then
                  strReceiver = strReceiver & vbCrLf & String(4, " ") & cu(j)
               End If
            Next j
            stLetter = Replace(stLetter, "TO: ", "TO: " & strReceiver, 1, 1)
         End If
         '2006/3/24 end
         
         strExc(0) = m_strFax(0)
         strExc(1) = m_strFax(1)
         'Add by Morgan 2009/2/16 若有傳真號碼2,電話號碼2時也要印
         If m_strFax(2) <> "" Then
            If m_strFax(0) <> "" Then
               strExc(0) = strExc(0) & ", " & m_strFax(2)
            Else
               strExc(0) = m_strFax(2)
            End If
         End If
         If m_strFax(3) <> "" Then
            If m_strFax(1) <> "" Then
               strExc(1) = strExc(1) & ", " & m_strFax(3)
            Else
               strExc(1) = m_strFax(3)
            End If
         End If
         'end 2009/2/16
         stLetter = Replace(stLetter, "FAX: ", "FAX: " & strExc(0), 1, 1)
         stLetter = Replace(stLetter, "TEL: ", "TEL: " & strExc(1), 1, 1)
         'Added by Morgan 2024/1/15
         If pa(1) = "CFP" And Option6.Value = True Then
            stLetter = Replace(stLetter, "E-Mail: ", "E-Mail: " & m_strCP44_FA16)
         Else
         'end 2024/1/15
      
            'Added by Morgan 2012/3/7 CFP 取消 TEL 改 E-Mail
            'Modify By Sindy 2014/9/18
            'stLetter = Replace(stLetter, "E-Mail: ", "E-Mail: " & GetEMail, 1, 1)
            'Modify By Sindy 2017/8/16 + , IIf(Option2.Value = True, "FC", IIf(Option6.Value = True, "CF", ""))
            stLetter = Replace(stLetter, "E-Mail: ", "E-Mail: " & PUB_GetFCeMailConText("Main_EMail", pa(1), pa(2), pa(3), pa(4), IIf(Option2.Value = True, "FC", IIf(Option6.Value = True, "CF", ""))), 1, 1)
            '2014/9/18 END
         End If
      End If
   End If
    
   bolRetry = False
    
On Error GoTo ERRORSECTION1
    
   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
   g_WordAp.Documents.add
   With g_WordAp
      'Add by Morgan 2013/1/8
      '切換為整頁模式,信頭才會正常顯示
      If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
         .ActiveWindow.ActivePane.View.Type = wdPageView
      Else
         .ActiveWindow.View.Type = wdPageView
      End If
      'end 2013/1/8
      
      '設定字型版面(參照定稿)
      .Selection.Font.Name = "Times New Roman"
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.Orientation = wdTextOrientationHorizontal
      .Selection.Font.Size = 12
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(3.175)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(3.175)
      'Modify by Morgan 2011/7/7 新信頭改版面
      '.Selection.PageSetup.TopMargin = .CentimetersToPoints(3.53)
      '.Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
      '.Selection.PageSetup.TopMargin = .CentimetersToPoints(5)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(4) 'Modify By Sindy 2015/10/14 外專要改開窗信封,因此為了看到地址姓名而調整
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(3)
      .Selection.PageSetup.FooterDistance = .CentimetersToPoints(2.5)
      'end 2011/7/7
      .Selection.ParagraphFormat.DisableLineHeightGrid = True
      '不要分散對齊
      '.Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
      
'Remove by Morgan 2011/7/12 因為第 2 頁以後不要有信頭故改回放在本文
'      'Add by Morgan 2011/7/6 新信頭改放在頁首頁尾
'      If txtLetterHead <> "N" Then
'         If PUB_ReadDB2File(stFileName, iPicNo) = True Then
'            .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader '改回
'            Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
'            oShape.ZOrder 4
'            oShape.LockAnchor = True
'            oShape.LockAspectRatio = -1
'            oShape.Width = .CentimetersToPoints(21)
'            oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
'            oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
'            oShape.Left = .CentimetersToPoints(0)
'            oShape.Top = .CentimetersToPoints(0)
'            If PUB_ReadDB2File(stFileName, iPicNo2) = True Then
'               .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
'               Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
'               oShape.ZOrder 4
'               oShape.LockAnchor = True
'               oShape.LockAspectRatio = -1
'               oShape.Width = .CentimetersToPoints(21)
'               oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
'               oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
'               oShape.Left = .CentimetersToPoints(0)
'               oShape.Top = .CentimetersToPoints(27.3)
'            End If
'            .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
'            .Selection.EndKey Unit:=wdStory
'         End If
'      End If
      
      
      '去掉前面的跳行符號
      If InStr(stLetter, Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10)) = 1 Then
         stLetter = Mid(stLetter, 9)
      End If
      'end 2011/7/6
      
      '印傳真封面
      If stLetter <> "" Then
         'Add by Morgan 2006/8/24
         '傳真封面信頭
         If txtLetterHead <> "N" Then
'Remove by Morgan 2011/7/6 新信頭改放在頁首頁尾
'            If PUB_ReadDB2File(stFileName, iPicNo) = True Then
'               '插入圖片檔案
'               .ActiveDocument.Shapes.AddPicture Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True
'               .ActiveDocument.Shapes("Picture " & Trim(.ActiveDocument.Shapes.Count + 1)).Select
'               .Selection.ShapeRange.ZOrder 4
'               .Selection.ShapeRange.LockAnchor = True
'               .Selection.ShapeRange.LockAspectRatio = -1
'               .Selection.ShapeRange.Width = 546.5
'               .Selection.ShapeRange.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
'               .Selection.ShapeRange.RelativeVerticalPosition = wdRelativeVerticalPositionPage
'               .Selection.ShapeRange.Left = .CentimetersToPoints(1)
'               .Selection.ShapeRange.Top = .CentimetersToPoints(1)
'               .Selection.ShapeRange.WrapFormat.Type = wdWrapNone 'Add by Morgan 2010/11/29
'               .Selection.EndKey Unit:=wdStory
'            End If
'            'end 2006/8/24

            'Add by Morgan 2011/7/12 因為第 2 頁以後不要有信頭故改回放在本文
            If PUB_ReadDB2File(stFileName, iPicNo) = True Then
               Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
               oShape.ZOrder 4
               oShape.LockAnchor = True
               oShape.LockAspectRatio = -1
               oShape.Width = .CentimetersToPoints(21)
               oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
               oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
               oShape.Left = .CentimetersToPoints(0)
               oShape.Top = .CentimetersToPoints(0)
               If PUB_ReadDB2File(stFileName, iPicNo2) = True Then
                  Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
                  oShape.ZOrder 4
                  oShape.LockAnchor = True
                  oShape.LockAspectRatio = -1
                  oShape.Width = .CentimetersToPoints(21)
                  oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                  oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                  oShape.Left = .CentimetersToPoints(0)
                  oShape.Top = .CentimetersToPoints(27.6)
               End If
               .Selection.EndKey Unit:=wdStory
            End If
         
         End If
         .Selection.TypeText stLetter
      End If
      
      'Add by Morgan 2006/8/24
      '信函信頭
      
'Remove by Morgan 2011/7/6 新信頭改放在頁首頁尾
'      If txtLetterHead <> "N" Then
'         If PUB_ReadDB2File(stFileName, iPicNo) = True Then
'            '插入圖片檔案
'            .ActiveDocument.Shapes.AddPicture Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True
'            .ActiveDocument.Shapes("Picture " & Trim(.ActiveDocument.Shapes.Count + 1)).Select
'            .Selection.ShapeRange.ZOrder 4
'            .Selection.ShapeRange.LockAnchor = True
'            .Selection.ShapeRange.LockAspectRatio = -1
'            .Selection.ShapeRange.Width = 546.5
'            .Selection.ShapeRange.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
'            .Selection.ShapeRange.RelativeVerticalPosition = wdRelativeVerticalPositionPage
'            .Selection.ShapeRange.Left = .CentimetersToPoints(1)
'            .Selection.ShapeRange.Top = .CentimetersToPoints(1)
'            .Selection.ShapeRange.WrapFormat.Type = wdWrapNone 'Add by Morgan 2010/11/29
'            .Selection.EndKey Unit:=wdStory
'         End If
'         'Add by Morgan 2010/11/29
'         .Selection.TypeParagraph
'         .Selection.TypeParagraph
'         .Selection.TypeParagraph
'         'end 2010/11/29
'      Else
'         .Selection.TypeParagraph
'         .Selection.TypeParagraph
'         .Selection.TypeParagraph
'      End If
'      'end 2006/8/24
'      .Selection.TypeParagraph
'end 2011/7/6

         'Add by Morgan 2011/7/12 因為第 2 頁以後不要有信頭故改回放在本文
         If txtLetterHead <> "N" Then
            If PUB_ReadDB2File(stFileName, iPicNo) = True Then
               Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
               oShape.ZOrder 4
               oShape.LockAnchor = True
               oShape.LockAspectRatio = -1
               oShape.Width = .CentimetersToPoints(21)
               oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
               oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
               oShape.Left = .CentimetersToPoints(0)
               oShape.Top = .CentimetersToPoints(0)
               If PUB_ReadDB2File(stFileName, iPicNo2) = True Then
                  Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
                  oShape.ZOrder 4
                  oShape.LockAnchor = True
                  oShape.LockAspectRatio = -1
                  oShape.Width = .CentimetersToPoints(21)
                  oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                  oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                  oShape.Left = .CentimetersToPoints(0)
                  oShape.Top = .CentimetersToPoints(27.6)
               End If
               .Selection.EndKey Unit:=wdStory
            End If
         End If
      

'Remove by Morgan 2006/9/20 傳真一定要有封面
'      '有印傳真封面時不需再印
'      If txtFaxFace = "N" Then
'         .Selection.TypeText "                                       By fax 1 Page(s) including this page"
'         .Selection.TypeParagraph
'      End If
'End 2006/9/20
        
      'Add by Morgan 2007/1/19 傳真號碼
      strExc(0) = m_strFax(0)
      If m_strFax(2) <> "" Then
         If m_strFax(0) <> "" Then
            strExc(0) = strExc(0) & ", " & m_strFax(2)
         Else
            strExc(0) = m_strFax(2)
         End If
      End If
      
      'Added by Morgan 2024/1/15
      If pa(1) = "CFP" And Option6.Value = True Then
         .Selection.TypeText "E-mail: " & m_strCP44_FA16
      Else
      'end 2024/1/15
      
         'edit by nickc 2007/08/21 去FAX跟: 中間的空白，Wayne 要求的
         '.Selection.TypeText "Fax : " & strExc(0)
         'Modify by Morgan 2008/4/8
         '.Selection.TypeText "Fax: " & strExc(0)
         If m_bEMail = True Then
            'Modify By Sindy 2014/9/18
            '.Selection.TypeText "E-mail: " & GetEMail
            'Modify By Sindy 2017/8/16 + , IIf(Option2.Value = True, "FC", IIf(Option6.Value = True, "CF", ""))
            .Selection.TypeText "E-mail: " & PUB_GetFCeMailConText("Main_EMail", pa(1), pa(2), pa(3), pa(4), IIf(Option2.Value = True, "FC", IIf(Option6.Value = True, "CF", "")))
            '2014/9/18 END
         Else
            .Selection.TypeText "Fax: " & strExc(0)
         End If
      End If
      .Selection.TypeParagraph
      'end 2007/1/19
      
      '系統日期(從中間開始印)
      '.Selection.TypeText Space(30) & ChgEngDate(strSrvDate(1))
      '置中
      If pa(1) = "CFP" Or pa(1) = "P" Then
         'Modified by Morgan 2013/8/20 改回從中間開始印但用定位點
         '.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .Selection.ParagraphFormat.FirstLineIndent = .CentimetersToPoints(7.35)
         .Selection.TypeText ChgEngDate(strSrvDate(1))
         .Selection.TypeParagraph
         .Selection.ParagraphFormat.FirstLineIndent = 0
      Else
         .Selection.TypeText Space(30) & ChgEngDate(strSrvDate(1))
      End If
      
      '靠左
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      'Modify By Sindy 2014/8/15 內文改寫Call共用Func
      .Selection.TypeText GetContentEnglish("發信對象", False)
      
      '發信對象為FC代理人
      'Modified by Lydia 2020/08/26 (限外專)受文者非X編號，不論是FC或CF代理人以及語言都帶備註出來；其他部門的都先不動，將來有需要再加。
      'If Option2.Value Then
      If Left(m_StrUserST03, 2) = "F2" And ((strSrvDate(1) >= 各項指示啟用日 And (Option2.Value = True Or Option6.Value = True)) Or (strSrvDate(1) < 各項指示啟用日 And Option2.Value = True)) Then
         PrintMemo 'Added by Morgan 2011/11/18
      End If
      'end 2020/08/26
      
      .Selection.TypeParagraph
      .Selection.TypeText GetContentEnglish("ATTN", False)
      .Selection.TypeParagraph
      .Selection.TypeText GetContentEnglish("RE", False)
      .Selection.TypeParagraph
      .Selection.TypeText GetContentEnglish("稱謂", False)
      .Selection.TypeParagraph
      'Modified by Morgan 2020/12/30
      '.Selection.TypeParagraph
      If m_IDSCP09 <> "" Then
         InsIDSContent m_IDSCP09
         bolShow = True 'Added by Morgan 2024/3/29
      'Added by Morgan 2021/2/2 寶齡富錦 Y55435 特殊控制
      ElseIf bIsBPFCase Then
         .Selection.TypeText "Please be informed that we received a "
         .Selection.Font.ColorIndex = wdRed
         .Selection.TypeText "官方來函性質"
         .Selection.Font.ColorIndex = wdAuto
         .Selection.TypeText ". Attached please find a copy of the Office Action."
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.TypeText "Please note that the due date for responding to the Office Action is "
         .Selection.Font.Bold = True
         .Selection.Font.Underline = wdUnderlineSingle
         If stNP09 <> "" Then
            .Selection.TypeText ChgEngDate(stNP09)
         Else
            .Selection.Font.ColorIndex = wdRed
            .Selection.TypeText "法定期限"
            .Selection.Font.ColorIndex = wdAuto
         End If
         .Selection.Font.Bold = False
         .Selection.Font.Underline = wdUnderlineNone
         .Selection.TypeText " and can be extended for two months upon request and payment of the necessary extension fees. We would appreciate receiving your instructions and remarks "
         .Selection.Font.Bold = True
         .Selection.Font.Underline = wdUnderlineSingle
         .Selection.TypeText "by our working deadline of "
         If stNP23 <> "" Then
            .Selection.TypeText ChgEngDate(stNP23)
         Else
            .Selection.Font.ColorIndex = wdRed
            .Selection.TypeText "約定期限"
            .Selection.Font.ColorIndex = wdAuto
         End If
         .Selection.TypeText " at the latest"
         .Selection.Font.Bold = False
         .Selection.Font.Underline = wdUnderlineNone
         .Selection.TypeText "."
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.TypeText "If you need us to prepare a summary English translation of the Office Action and/or our comments, please notify us as soon as possible."
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.TypeText "We look forward to receiving your timely instruction."
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.TypeText "Should you have any questions concerning the referenced application, please do not hesitate to contact us. Thank you."
         .Selection.TypeParagraph
      'end 2021/2/2
      End If
      'end 2020/12/30
      '2014/8/15 END
      
      'Added by Morgan 2024/3/29
      'FCP二次核對報告
      If pa(1) = "FCP" And Val(Right(Combo8, 4)) = "926" Then
         'Modified by Morgan 2024/4/12 取消非設計案限制
         If pa(173) = "" Then MsgBox "「系統無圖式資訊」", vbExclamation
         strExc(1) = PUB_Num2DocNo(Combo8.ItemData(Combo8.ListIndex))
         strExc(8) = ""
         NowPrint strExc(1), "07", "01", False, strUserNum, , , True, strExc(8)
         If strExc(8) <> "" Then
            .Selection.TypeText strExc(8)
         End If
         bolShow = True
      'Add by Morgan 2006/9/20 -- David
      ElseIf Left(m_Dept, 2) = "F2" Or pa(1) = "FCP" Or pa(1) = "FG" Then
         .Selection.TypeParagraph
         'Modify by Morgan 2007/3/5 --David
         '.Selection.TypeText "     With best regards,"
'         .Selection.TypeText "With best regards,"
         'end 2007/3/5
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         'Modify By Sindy 2013/10/29
         '.Selection.TypeText Space(30) & "Sincerely yours,"
         .Selection.TypeText Space(30) & "Best regards,"
         '2013/10/29 END
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.TypeText Space(30) & "Fred C. T. Yen"
         .Selection.TypeParagraph
         'Modified by Morgan 2012/3/20
         '.Selection.TypeText Space(30) & "Patent Attorney"
         .Selection.TypeText "Patent Department" & Space(15) & "Patent Attorney"
         .Selection.TypeParagraph
         'Modified by Morgan 2014/2/11
         If m_CompNo = "J" Then
            .Selection.TypeText Space(30) & "Tai E Intellectual Property Co., Ltd."
         Else
            .Selection.TypeText Space(30) & "Tai E International Patent & Law Office"
         End If
         'end 2014/2/11
         .Selection.TypeParagraph
         .Selection.TypeText "CTY/" & Pub_StrUserSt17
      
      'Add by Morgan 2007/2/13 -- Amy
      ElseIf pa(1) = "CFP" Or pa(1) = "P" Then
         .Selection.TypeParagraph
         'Modify by Morgan 2007/8/29 --Amy請作單
         '.Selection.TypeText "With best regards,"
         '.Selection.TypeParagraph
         '.Selection.TypeParagraph
         '.Selection.TypeText Space(30) & "Sincerely yours,"
         '.Selection.TypeParagraph
         '.Selection.TypeParagraph
         '.Selection.TypeParagraph
         '.Selection.TypeParagraph
         '.Selection.TypeText Space(30) & "Fred C. T. Yen"
         '.Selection.TypeParagraph
         '.Selection.TypeText Space(30) & "Patent Attorney"
         '.Selection.TypeParagraph
         '.Selection.TypeText Space(30) & "Tai E International Patent & Law Office"
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         'Modified by Morgan 2013/8/20
         '.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .Selection.ParagraphFormat.FirstLineIndent = .CentimetersToPoints(7.35)
         'Modify By Sindy 2013/10/29
         '.Selection.TypeText "Sincerely yours,"
         .Selection.TypeText "Best regards,"
         '2013/10/29 END
         .Selection.TypeParagraph
         'Modified by Morgan 2015/8/18
         '.Selection.TypeParagraph
         '.Selection.TypeParagraph
         .Selection.TypeText "|#(林景郁英文自動簽名檔)#|"
         'end 2015/8/18
         .Selection.TypeParagraph
         'Modify by Morgan 2010/5/10
         '.Selection.TypeText "Fred C. T. Yen"
         .Selection.TypeText "Jerry C. Y. Lin"
         'end 2010/5/10
         .Selection.TypeParagraph
         .Selection.TypeText "Patent Attorney"
         .Selection.TypeParagraph
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
         'end 2007/8/29
         .Selection.TypeParagraph
         .Selection.ParagraphFormat.FirstLineIndent = 0 'Added by Morgan 2013/8/20
         
         'Modify by Morgan 2010/5/10
         '.Selection.TypeText "CTY/" & Pub_StrUserSt17
         .Selection.TypeText "CYL/" & Pub_StrUserSt17
         'end 2010/5/10
      Else
         .Selection.TypeParagraph
         'Modify by Morgan 2007/3/5 --David
         '.Selection.TypeText "     With best regards,"
'         .Selection.TypeText "With best regards,"
         'end 2007/3/5
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         'Modified by Morgan 2013/8/20 與日期對齊
         '.Selection.TypeText "　　　　　　　　　　　　　　　　　　　Sincerely yours,"
         'Modify By Sindy 2013/10/29
         '.Selection.TypeText Space(30) & "Sincerely yours,"
         .Selection.TypeText Space(30) & "Best regards,"
         '2013/10/29 END
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         .Selection.TypeParagraph
         'Modified by Morgan 2013/8/20 與日期對齊
         '.Selection.TypeText "　　　　　　　　　　　　　　　　　　　Tai E International"
         
         'Modified by Morgan 2014/2/11
         '.Selection.TypeText Space(30) & "Tai E International"
         If m_CompNo = "J" Then
            .Selection.TypeText Space(30) & "Tai E Intellectual"
         Else
            .Selection.TypeText Space(30) & "Tai E International"
         End If
         'end 2014/2/11
         
         .Selection.TypeParagraph
         'Modified by Morgan 2013/8/20 與日期對齊
         '.Selection.TypeText "　　　　　　　　　　　　　　　　　　　Patent & Law Office"
         
         'Modified by Morgan 2014/2/11
         '.Selection.TypeText Space(30) & "Patent & Law Office"
         If m_CompNo = "J" Then
            .Selection.TypeText Space(30) & "Property Co., Ltd."
         Else
            .Selection.TypeText Space(30) & "Patent & Law Office"
         End If
         'end 2014/2/11
         
         .Selection.TypeParagraph
         .Selection.TypeText "CTY/" & Pub_StrUserSt17
      End If
      
      
      '操作程序專業代號-對國外
'Modify by Morgan 2007/1/25 改抓全域變數
'      strExc(0) = "SELECT ST17 FROM STAFF WHERE ST01='" & strUserNum & "'"
'      intI = 1
'      Set rsTemp = clslawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         If IsNull(rsTemp.Fields(0).Value) Then
'            'Modify by Morgan 2006/7/6
'            '.Selection.TypeText "ICL/"
'            .Selection.TypeText "CTY/"
'         Else
'            'Modify by Morgan 2006/7/6
'            '.Selection.TypeText "ICL/" & rsTemp.Fields(0).Value
'            .Selection.TypeText "CTY/" & rsTemp.Fields(0).Value
'         End If
'      End If
      
      'Remove by Morgan 2010/5/7 改依部門別不同
      '.Selection.TypeText "CTY/" & Pub_StrUserSt17
      
      'Add by Morgan 2008/12/4
      If pa(1) = "CFP" Then
         'Added by Morgan 2023/9/12
         If ChkBCP() = True Then
            .Selection.TypeParagraph
            .Selection.TypeParagraph
            .Selection.TypeText "*若此程序為B類內部收文即表示為本所疏失或代理人疏失不會向客戶收取費用，則請依案情在指示信中帶入以下的內容："
            .Selection.TypeParagraph
            .Selection.TypeText "(1)本所疏失－"
            .Selection.TypeParagraph
            .Selection.TypeText "This correction/amendment is to remedy our office's oversight and the client will not be billed; we would like to request you to kindly offer a most-favorable discount on the service fee."
            .Selection.TypeParagraph
            .Selection.TypeText "(2)代理人疏失－"
            .Selection.TypeParagraph
            .Selection.TypeText "This correction/amendment is to remedy your office's oversight and we are unable to charge the client for this procedure. We would like to request you to absorb the expenses incurred therefrom. Thank you for your assistance."
         End If
         'end 2023/9/12
         
         'Modified by Lydia 2019/09/27 改成共用
         'strExc(0) = GetLetterMemo(pa(1), "2")
         strExc(0) = Pub_GetLetterMemo(pa(1), "2")
         If strExc(0) <> "" Then
            .Selection.TypeParagraph
            .Selection.TypeParagraph
            .Selection.TypeText strExc(0)
         End If
      End If
'end 2007/1/25
      .Selection.WholeStory
      ChgWordFormat g_WordAp, .Selection.Text
   End With
    
   g_WordAp.Visible = True
   'Added by Morgan 2024/3/26
   If bolShow Then
      g_WordAp.Selection.HomeKey Unit:=wdStory
      g_WordAp.Activate
   Else
   'end 2024/3/26
      g_WordAp.WindowState = wdWindowStateMaximize
   End If
   Set g_WordAp = Nothing 'Added by Morgan 2015/9/7
   
ERRORSECTION1:
   If Err.NUMBER <> 0 Then
      Select Case Err.NUMBER
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

Private Sub Text5_GotFocus()
    TextInverse Text5
End Sub

Private Sub Text6_GotFocus(Index As Integer)
    TextInverse Text6(Index)
End Sub

Private Sub txtFaxFace_GotFocus()
   TextInverse txtFaxFace
End Sub

Private Sub txtFaxFace_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> 78 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtLetterHead_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
      KeyAscii = 0
   End If
End Sub

Private Function GetNPSQL() As String
   'Modofied by Morgan 2021/2/2 +np09
   'Modified by Morgan 2021/8/2 +np23
   strExc(0) = "SELECT NP08,NP09,NP23 FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
   " AND NP06 IS NULL"
   'Add by Morgan 2007/1/3 排除下列程序管制的性質
   '2010/3/22 MODIFY BY SONIA 改以strNpSqlOfNoSalesDuty控制
   strExc(0) = strExc(0) & strNpSqlOfNoSalesDuty
   'end 2007/1/3
   
   strExc(0) = strExc(0) & " order by np08 asc" 'Add by Morgan 2007/11/1 抓期限最近的
   
   GetNPSQL = strExc(0)
End Function

'日文信
Private Sub WordJapan()
   Dim stLetter As String
   Dim m As Integer, j As Integer
   Dim strCustName As String
   Dim strReceiver As String
   Dim stFileName As String '暫存圖檔檔名
   Dim iPicNo As Integer, iPicNo2 As Integer '上,下圖檔代碼
   Dim oShape
   Dim stAppCountry As String '申請國家
   Dim StrSQLa As String 'Added by Lydia 2022/03/29
   Dim strUniText As String 'Added by Morgan 2022/7/27
   
'Added by Morgan 2020/3/25
If strSrvDate(1) >= 智慧所更名日 Then
   m_CompNo = PUB_GetReceiptComp(pa(1), pa(2), pa(3), pa(4), True)
   PUB_GetLetterPicID m_CompNo, pa(1), iPicNo, iPicNo2, 3, True, m_Dept
Else
'end 2020/3/25
   
'Modify by Morgan 2011/7/6 改用新的信頭
'   Select Case Text1.Text
'      Case "FCT", "CFT"
'         iPicNo = 1
'      'Modify By Sindy 2009/07/24 增加LIN系統類別
'      Case "FCP", "FG", "FCL", "CFL", "LIN"
'         iPicNo = 2
'      Case "CFP", "CPS"
'         iPicNo = 3
'      Case Else
'         iPicNo = 4
'   End Select
   iPicNo = 5
   iPicNo2 = 9
'end 2011/7/6

   'Added by Morgan 2013/12/30
   m_CompNo = PUB_GetReceiptComp(pa(1), pa(2), pa(3), pa(4))
   '智權公司
   If m_CompNo = "J" Then
      iPicNo = 21
      'patent信箱
      If pa(1) = "CFP" Then
         iPicNo2 = 22
      'ipdept信箱
      ElseIf pa(1) = "CFT" Then
         iPicNo2 = 24
      End If
   '專利處對國外都用專利法律公司(patent信箱)
   ElseIf Left(m_Dept, 1) <> "F" And (pa(1) = "CFP" Or pa(1) = "P") Then
      iPicNo2 = 18
   End If
   'end 2013/12/30
   
End If 'Added by Morgan 2020/3/25
         
'Modified by Morgan 2013/11/22 不管是否印傳真封面都要抓發信對象

      '因撰寫信函傳本所號&000,所以進度檔相關欄位不會有資料
         If Option2.Value = True Then
            If fa(5) <> "" Then
               strReceiver = fa(5)
            ElseIf fa(1) <> "" Then
               strReceiver = Trim(fa(1))
               If fa(2) <> "" Then
                  strReceiver = strReceiver & " " & Trim(fa(2))
               End If
               If fa(3) <> "" Then
                  strReceiver = strReceiver & vbCrLf & Space(8) & Trim(fa(3))
               End If
               If fa(4) <> "" Then
                  strReceiver = strReceiver & " " & Trim(fa(4))
               End If
            Else
               strReceiver = fa(0)
            End If
         ElseIf Option6.Value = True Then
            If cfa(5) <> "" Then
               strReceiver = cfa(5)
            ElseIf cfa(1) <> "" Then
               strReceiver = Trim(cfa(1))
               If cfa(2) <> "" Then
                  strReceiver = strReceiver & " " & Trim(cfa(2))
               End If
               If cfa(3) <> "" Then
                  strReceiver = strReceiver & vbCrLf & Space(8) & Trim(cfa(3))
               End If
               If cfa(4) <> "" Then
                  strReceiver = strReceiver & " " & Trim(cfa(4))
               End If
            Else
               strReceiver = cfa(0)
            End If
         ElseIf Option3.Value = True Then
            If cu(5) <> "" Then
               strReceiver = cu(5)
            ElseIf cu(1) <> "" Then
               strReceiver = Trim(cu(1))
               If cu(2) <> "" Then
                  strReceiver = strReceiver & " " & Trim(cu(2))
               End If
               If cu(3) <> "" Then
                  strReceiver = strReceiver & vbCrLf & Space(8) & Trim(cu(3))
               End If
               If cu(4) <> "" Then
                  strReceiver = strReceiver & " " & Trim(cu(4))
               End If
            Else
               strReceiver = cu(0)
            End If
         End If
         
'end 2013/11/22
         
   '是否印傳真封面
   If txtFaxFace <> "N" Then
      intI = vbYes
      If m_bEMail = True Then
         intI = MsgBox("本案有設定以EMail通知，是否要印傳真封面?", vbYesNo + vbDefaultButton2)
      End If
      If intI = vbYes Then
      
         Select Case Text1.Text
            Case "FCP"
               NowPrint Text1.Text & Text2.Text & Text3.Text & Text4.Text & "&000", "04", "Z9", False, strUserNum, , , True, stLetter, 1, Label11.Caption
            Case Else
               'Add by Morgan 2010/9/7
               If Left(Text1, 1) = "C" And m_strCP09 <> "" Then
                  NowPrint m_strCP09, "01", "Z9", False, strUserNum, , , True, stLetter, 1, Label11.Caption
               Else
               'end 2010/9/7
                  NowPrint Text1.Text & Text2.Text & Text3.Text & Text4.Text & "&000", "01", "Z9", False, strUserNum, , , True, stLetter, 1, Label11.Caption
               End If
         End Select
         If stLetter <> "" Then stLetter = stLetter & Chr(12)
         
         stLetter = Replace(stLetter, "受信者：", "受信者：" & strReceiver, 1, 1)
         'modify by Morgan 2014/4/23 若有傳真號碼2,電話號碼2時也要印
         'stLetter = Replace(stLetter, "иャЧヱЗ番�A：", "иャЧヱЗ番�A：" & m_strFax(0), 1, 1)
         'stLetter = Replace(stLetter, "電　話　番　�A：", "電　話　番　�A：" & m_strFax(1), 1, 1)
         strExc(0) = m_strFax(0)
         strExc(1) = m_strFax(1)
         If m_strFax(2) <> "" Then
            If m_strFax(0) <> "" Then
               strExc(0) = strExc(0) & ", " & m_strFax(2)
            Else
               strExc(0) = m_strFax(2)
            End If
         End If
         If m_strFax(3) <> "" Then
            If m_strFax(1) <> "" Then
               strExc(1) = strExc(1) & ", " & m_strFax(3)
            Else
               strExc(1) = m_strFax(3)
            End If
         End If
         'Modified by Morgan 2022/7/27
         'stLetter = Replace(stLetter, "иャЧヱЗ番�A：", "иャЧヱЗ番�A：" & strExc(0), 1, 1)
         'stLetter = Replace(stLetter, "電　話　番　�A：", "電　話　番　�A：" & strExc(1), 1, 1)
         strUniText = PUB_GetUniText(Me.Name, "傳真")
         stLetter = Replace(stLetter, strUniText, strUniText & strExc(0), 1, 1)
         strUniText = PUB_GetUniText(Me.Name, "電話")
         stLetter = Replace(stLetter, strUniText, strUniText & strExc(1), 1, 1)
         'end 2022/7/27
         '2014/4/23 end
      End If
   End If
    
   bolRetry = False
    
On Error GoTo ERRORSECTION1
    
   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
   g_WordAp.Documents.add
   With g_WordAp
      'Add by Morgan 2013/1/8
      '切換為整頁模式,信頭才會正常顯示
      If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
         .ActiveWindow.ActivePane.View.Type = wdPageView
      Else
         .ActiveWindow.View.Type = wdPageView
      End If
      'end 2013/1/8
      
      '設定字型版面(參照定稿)
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.Orientation = wdTextOrientationHorizontal
      'Modify by Morgan 2011/7/6
      .Selection.Font.Name = "MS Mincho" 'Added by Morgan 2022/7/13 改 MS Mincho -- 毓芳, May
      If pa(1) = "FCP" Then
         'Modified by Morgan 2014/9/24 --毓芳
         '.Selection.Font.Name = "標楷體"
         '.Selection.Font.Name = "細明體"
         'end 2014/9/24
         .Selection.Font.Size = 12
      ElseIf pa(1) = "FCT" Or pa(1) = "CFT" Or pa(1) = "CFC" Then
         '.Selection.Font.Name = "細明體"
         .Selection.Font.Size = 13
      Else
         '.Selection.Font.Name = "細明體"
         .Selection.Font.Size = 12
      End If
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(3.175)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(3.175)
      'Modify by Morgan 2011/7/7 新信頭改版面
      '.Selection.PageSetup.TopMargin = .CentimetersToPoints(3.53)
      '.Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
      '.Selection.PageSetup.TopMargin = .CentimetersToPoints(5)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(4) 'Modify By Sindy 2015/10/14 外專要改開窗信封,因此為了看到地址姓名而調整
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(3)
      .Selection.PageSetup.FooterDistance = .CentimetersToPoints(2.5)
      'end 2011/7/7
      .Selection.ParagraphFormat.DisableLineHeightGrid = True
      
'Remove by Morgan 2011/7/12 因為第 2 頁以後不要有信頭故改回放在本文
'      'Add by Morgan 2011/7/6 新信頭改放在頁首頁尾
'      If txtLetterHead <> "N" Then
'         If PUB_ReadDB2File(stFileName, iPicNo) = True Then
'            .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
'            Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
'            oShape.ZOrder 4
'            oShape.LockAnchor = True
'            oShape.LockAspectRatio = -1
'            oShape.Width = .CentimetersToPoints(21)
'            oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
'            oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
'            oShape.Left = .CentimetersToPoints(0)
'            oShape.Top = .CentimetersToPoints(0)
'            If PUB_ReadDB2File(stFileName, iPicNo2) = True Then
'               .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
'               Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
'               oShape.ZOrder 4
'               oShape.LockAnchor = True
'               oShape.LockAspectRatio = -1
'               oShape.Width = .CentimetersToPoints(21)
'               oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
'               oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
'               oShape.Left = .CentimetersToPoints(0)
'               oShape.Top = .CentimetersToPoints(27.3)
'            End If
'            .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
'            .Selection.EndKey Unit:=wdStory
'         End If
'      End If
      
      '去掉前面的跳行符號
      If InStr(stLetter, Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10)) = 1 Then
         stLetter = Mid(stLetter, 9)
      End If
      'end 2011/7/6
      
      '印傳真封面
      If stLetter <> "" Then
         'Add by Morgan 2006/8/24
         '傳真封面信頭
         If txtLetterHead <> "N" Then
'Remove by Morgan 2011/7/6 新信頭改放在頁首頁尾
'            If PUB_ReadDB2File(stFileName, iPicNo) = True Then
'               '插入圖片檔案
'               .ActiveDocument.Shapes.AddPicture Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True
'               .ActiveDocument.Shapes("Picture " & Trim(.ActiveDocument.Shapes.Count + 1)).Select
'               .Selection.ShapeRange.ZOrder 4
'               .Selection.ShapeRange.LockAnchor = True
'               .Selection.ShapeRange.LockAspectRatio = -1
'               .Selection.ShapeRange.Width = 546.5
'               .Selection.ShapeRange.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
'               .Selection.ShapeRange.RelativeVerticalPosition = wdRelativeVerticalPositionPage
'               .Selection.ShapeRange.Left = .CentimetersToPoints(1)
'               .Selection.ShapeRange.Top = .CentimetersToPoints(1)
'               .Selection.ShapeRange.WrapFormat.Type = wdWrapNone 'Add by Morgan 2010/11/29
'               .Selection.EndKey Unit:=wdStory
'            End If
'            'end 2006/8/24

            'Add by Morgan 2011/7/12 因為第 2 頁以後不要有信頭故改回放在本文
            If PUB_ReadDB2File(stFileName, iPicNo) = True Then
               Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
               oShape.ZOrder 4
               oShape.LockAnchor = True
               oShape.LockAspectRatio = -1
               oShape.Width = .CentimetersToPoints(21)
               oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
               oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
               oShape.Left = .CentimetersToPoints(0)
               oShape.Top = .CentimetersToPoints(0)
               If PUB_ReadDB2File(stFileName, iPicNo2) = True Then
                  Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
                  oShape.ZOrder 4
                  oShape.LockAnchor = True
                  oShape.LockAspectRatio = -1
                  oShape.Width = .CentimetersToPoints(21)
                  oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                  oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                  oShape.Left = .CentimetersToPoints(0)
                  oShape.Top = .CentimetersToPoints(27.6)
               End If
               .Selection.EndKey Unit:=wdStory
            End If
         
         End If

         .Selection.TypeText stLetter
      End If
      
'Remove by Morgan 2011/7/6 新信頭改放在頁首頁尾
'      '信函信頭
'      If txtLetterHead <> "N" Then
'         If PUB_ReadDB2File(stFileName, iPicNo) = True Then
'            '插入圖片檔案
'            .ActiveDocument.Shapes.AddPicture Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True
'            .ActiveDocument.Shapes("Picture " & Trim(.ActiveDocument.Shapes.Count + 1)).Select
'            .Selection.ShapeRange.ZOrder 4
'            .Selection.ShapeRange.LockAnchor = True
'            .Selection.ShapeRange.LockAspectRatio = -1
'            .Selection.ShapeRange.Width = 546.5
'            .Selection.ShapeRange.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
'            .Selection.ShapeRange.RelativeVerticalPosition = wdRelativeVerticalPositionPage
'            .Selection.ShapeRange.Left = .CentimetersToPoints(1)
'            .Selection.ShapeRange.Top = .CentimetersToPoints(1)
'            .Selection.ShapeRange.WrapFormat.Type = wdWrapNone 'Add by Morgan 2010/11/29
'            .Selection.EndKey Unit:=wdStory
'         End If
'         'Add by Morgan 2010/11/29
'         .Selection.TypeParagraph
'         .Selection.TypeParagraph
'         .Selection.TypeParagraph
'         'end 2010/11/29
'      Else
'         .Selection.TypeParagraph
'         .Selection.TypeParagraph
'         .Selection.TypeParagraph
'      End If
'      .Selection.TypeParagraph
'end 2011/7/6
        
         'Add by Morgan 2011/7/12 因為第 2 頁以後不要有信頭故改回放在本文
         If txtLetterHead <> "N" Then
            If PUB_ReadDB2File(stFileName, iPicNo) = True Then
               Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
               oShape.ZOrder 4
               oShape.LockAnchor = True
               oShape.LockAspectRatio = -1
               oShape.Width = .CentimetersToPoints(21)
               oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
               oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
               oShape.Left = .CentimetersToPoints(0)
               oShape.Top = .CentimetersToPoints(0)
               If PUB_ReadDB2File(stFileName, iPicNo2) = True Then
                  Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
                  oShape.ZOrder 4
                  oShape.LockAnchor = True
                  oShape.LockAspectRatio = -1
                  oShape.Width = .CentimetersToPoints(21)
                  oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                  oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                  oShape.Left = .CentimetersToPoints(0)
                  oShape.Top = .CentimetersToPoints(27.6)
               End If
               .Selection.EndKey Unit:=wdStory
            End If
         End If
         
      strExc(0) = m_strFax(0)
      If m_strFax(2) <> "" Then
         If m_strFax(0) <> "" Then
            strExc(0) = strExc(0) & ", " & m_strFax(2)
         Else
            strExc(0) = m_strFax(2)
         End If
      End If
      
      '靠左
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      
      'Added by Morgan 2024/1/15
      If pa(1) = "CFP" And Option6.Value = True Then
         .Selection.TypeText "ＥＭＡＩＬ：" & m_strCP44_FA16
      Else
      'end 2024/1/15
      
         'Add by Morgan 2008/4/8
         If m_bEMail = True Then
            'Modify By Sindy 2014/9/18
            '.Selection.TypeText "ＥＭＡＩＬ：" & GetEMail
            'Modify By Sindy 2017/8/16 + , IIf(Option2.Value = True, "FC", IIf(Option6.Value = True, "CF", ""))
            .Selection.TypeText "ＥＭＡＩＬ：" & PUB_GetFCeMailConText("Main_EMail", pa(1), pa(2), pa(3), pa(4), IIf(Option2.Value = True, "FC", IIf(Option6.Value = True, "CF", "")))
            '2014/9/18 END
         Else
         'end 2008/4/8
            '傳真
            .Selection.TypeText "ＦＡＸ：" & strExc(0)
         End If
      End If
      
      .Selection.TypeParagraph
      
      '日期
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      .Selection.TypeText ExceptFieldData("系統日/中西")
      .Selection.TypeParagraph
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      .Selection.TypeParagraph
      
      '信函抬頭
      'Modified by Morgan 2013/11/22
      '沒有聯絡人才要帶"御中"
      If (Option4 And Trim(txtConDepJP & Combo4 & Combo5) = "") Or (Option5 And Trim(Combo6) = "") Then
         strReceiver = strReceiver & "　御中"
      End If
      'end 2013/11/22
      
      strReceiver = Replace(strReceiver, vbCrLf & Space(8), vbCrLf) '去掉前面空白
      .Selection.TypeText strReceiver
      .Selection.TypeParagraph
      'Modified by Lydia 2020/08/26 (限外專)受文者非X編號，不論是FC或CF代理人以及語言都帶備註出來；其他部門的都先不動，將來有需要再加。
      'If Option2.Value Then PrintMemo 'Added by Morgan 2011/11/18
      If Left(m_StrUserST03, 2) = "F2" And ((strSrvDate(1) >= 各項指示啟用日 And (Option2.Value = True Or Option6.Value = True)) Or (strSrvDate(1) < 各項指示啟用日 And Option2.Value = True)) Then
         PrintMemo
      End If
      'end 2020/08/26
      
      '聯絡人
      If Option4.Value Then
         If txtConDepJP <> "" Then
            .Selection.TypeText txtConDepJP
            .Selection.TypeParagraph
         End If
         If Combo4 <> "" Then
            .Selection.TypeText Combo4
            .Selection.TypeParagraph
         End If
         If Combo5 <> "" Then
            .Selection.TypeText Combo5
            .Selection.TypeParagraph
         End If
      ElseIf Option5.Value Then
         If Combo6 <> "" Then
            .Selection.TypeText Combo6
            .Selection.TypeParagraph
         End If
      End If
      
      If .Selection.Font.Size = 13 Then
         strExc(1) = String(17, "　")
      Else
         strExc(1) = String(19, "　")
      End If
      
      .Selection.TypeParagraph
      'Modified by Morgan 2020/4/1
      '.Selection.TypeText strExc(1) & "台一�罈��G利 (特許) 法律事務所"
      .Selection.TypeText strExc(1) & CompNameQuery("2", 3)
      'end 2020/4/1
      .Selection.TypeParagraph
      'Modified by Morgan 2022/7/27
      '.Selection.TypeText strExc(1) & "台北市長安東路二段112�A9階"
      strUniText = PUB_GetUniText(Me.Name, "地址")
      .Selection.TypeText strExc(1) & strUniText
      'end 2022/7/27
      .Selection.TypeParagraph
      'Modify by Morgan 2010/5/12
      If Text1.Text = "CFP" Then
         .Selection.TypeText strExc(1) & "弁理士　林 景 郁"
         .Selection.TypeParagraph
         'Modified by Morgan 2015/8/18
         .Selection.TypeText strExc(1) & "　　　　|#(林景郁中文自動簽名檔)#|"
         'end 2015/8/18
         .Selection.TypeParagraph
      Else
         'Modified by Lydia 2022/03/29
         '.Selection.TypeText strExc(1) & "弁理士　閻 啟 泰"
         StrSQLa = GetStaffName("81040", True)
         .Selection.TypeText strExc(1) & "弁理士　" & Mid(StrSQLa, 1, 1) & " " & Mid(StrSQLa, 2, 1) & " " & Mid(StrSQLa, 3, 1)
         'end 2022/03/29
         .Selection.TypeParagraph
         .Selection.TypeParagraph
      End If
      
      strExc(8) = ""
      'Modified by Morgan 2022/7/27
      'stAppCountry = "台��"
      stAppCountry = PUB_GetUniText(Me.Name, "台灣")
      'end 2022/7/27
      
      'Add by Morgan 2010/4/12 CFP案也會用非台灣則改抓申請國
      If Label11 <> "000" Then stAppCountry = Label12
      
      
      If pa(1) = "P" Or pa(1) = "CFP" Or pa(1) = "FCP" Then
         strExc(9) = ""
         If pa(7) <> "" Then
            strExc(9) = pa(7)
         ElseIf pa(6) <> "" Then
            strExc(9) = pa(6)
         Else
            strExc(9) = pa(5)
         End If
         .Selection.TypeText "件名：" & strExc(9)
         .Selection.TypeParagraph
         'Modified by Morgan 2022/7/27
         '.Selection.TypeText "　　　" & stAppCountry & GetPatentName(pa(8), 3) & "出願番�A：" & pa(11)
         strUniText = PUB_GetUniText(Me.Name, "申請號")
         .Selection.TypeText "　　　" & stAppCountry & GetPatentName(pa(8), 3) & strUniText & pa(11)
         'end 2022/7/27
         .Selection.TypeParagraph
         
         'Modify by Morgan 2011/8/11 台灣案公告號(日)才會等同證書號(日)
         If pa(1) = "FCP" Or (pa(1) = "P" And Label11 = "000") Then
            'Modified by Morgan 2022/7/27
            '.Selection.TypeText "　　　登�鰽f�A(公告番�A)：" & pa(22)
            strUniText = PUB_GetUniText(Me.Name, "公告號")
            .Selection.TypeText "　　　" & strUniText & pa(22)
            'end 2022/7/27
            .Selection.TypeParagraph
            'Modified by Morgan 2022/7/27
            '.Selection.TypeText "　　　登�髐�(公告日)：" & TranslateKeyWord(incCNV_CHINESE_CUN1, DBDATE(pa(21)), "")
            strUniText = PUB_GetUniText(Me.Name, "公告日")
            .Selection.TypeText "　　　" & strUniText & TranslateKeyWord(incCNV_CHINESE_CUN1, DBDATE(pa(21)), "")
            'end 2022/7/27
            .Selection.TypeParagraph
         Else
            'Modified by Morgan 2022/7/27
            '.Selection.TypeText "　　　登�鰽f�A：" & pa(22)
            strUniText = PUB_GetUniText(Me.Name, "證書號")
            .Selection.TypeText "　　　" & strUniText & pa(22)
            'end 2022/7/27
            .Selection.TypeParagraph
            
            'modify by sonia 2020/2/20 TranslateKeyWord(incCNV_CHINESE_MINKO...加傳本所案號,以判斷日期欄之民國或西元格式
            'Modified by Morgan 2022/7/27
            '.Selection.TypeText "　　　登�髐憿G" & TranslateKeyWord(incCNV_CHINESE_CUN, DBDATE(pa(21)), "", pa(1) & pa(2) & pa(3) & pa(4))
            strUniText = PUB_GetUniText(Me.Name, "發證日")
            .Selection.TypeText "　　　" & strUniText & TranslateKeyWord(incCNV_CHINESE_CUN, DBDATE(pa(21)), "", pa(1) & pa(2) & pa(3) & pa(4))
            'end 2022/7/27
            .Selection.TypeParagraph
         End If
         
         j = 0
      ElseIf pa(1) = "FCT" Then
         'Modified by Morgan 2022/7/27
         '.Selection.TypeText "件名：台�灠蚍迮n�鬙X願No. " & pa(12)
         strUniText = PUB_GetUniText(Me.Name, "件名")
         .Selection.TypeText strUniText & pa(12)
         'end 2022/7/27
         .Selection.TypeParagraph
         strExc(8) = "　　　"
         j = 51
      Else
         strExc(9) = ""
         If pa(22) <> "" Then
            strExc(9) = pa(22)
         ElseIf pa(6) <> "" Then
            strExc(9) = pa(6)
         Else
            strExc(9) = pa(5)
         End If
         strExc(8) = GetTradeMarkName(pa(8), 3)
         .Selection.TypeText "件名：" & stAppCountry & strExc(8) & "出願第 " & strExc(9) & " 號"
         
         .Selection.TypeParagraph
           
         .Selection.TypeText "　　　" & stAppCountry & strExc(8) & "出願  件"
         .Selection.TypeParagraph
         strExc(9) = ""
         If pa(7) <> "" Then
            strExc(9) = pa(7)
         ElseIf pa(6) <> "" Then
            strExc(9) = pa(6)
         Else
            strExc(9) = pa(5)
         End If
         'Modified by Morgan 2022/7/27
         '.Selection.TypeText "　　　名稱:" & strExc(9)
         strUniText = PUB_GetUniText(Me.Name, "名稱")
         .Selection.TypeText "　　　" & strUniText & strExc(9)
         'end 2022/7/27
         .Selection.TypeParagraph
         j = 51
      End If
      
      '申請人1
      strExc(9) = ""
      If cu(5) <> "" Then
         strExc(9) = cu(5)
      ElseIf cu(1) <> "" Then
         strExc(9) = Trim(cu(1)) & " " & Trim(cu(2))
         If cu(3) <> "" Then
            strExc(9) = strExc(9) & vbCrLf & String(4, "　") & Trim(cu(3)) & " " & Trim(cu(4))
         End If
      Else
         strExc(9) = cu(0)
      End If
      .Selection.TypeText strExc(8) & "　　　出願人：" & strExc(9)
      .Selection.TypeParagraph
      
      '申請人2~5
      For m = j + 27 To j + 30
         If pa(m) <> "" Then
            pa(m) = ChangeCustomerL(pa(m))
            strExc(0) = "select cu04,cu05,cu88,cu89,cu90,cu06 from customer where cu01||cu02='" & pa(m) & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strExc(9) = ""
               With RsTemp
                  If "" & .Fields("cu06") <> "" Then
                     strExc(9) = .Fields("cu06")
                  ElseIf "" & .Fields("cu05") <> "" Then
                     strExc(9) = Trim(.Fields("cu05")) & " " & Trim("" & .Fields("cu88"))
                     If "" & .Fields("cu89") <> "" Then
                        strExc(9) = strExc(9) & vbCrLf & String(4, "　") & Trim(.Fields("cu89")) & " " & Trim("" & .Fields("cu90"))
                     End If
                  Else
                     strExc(9) = "" & .Fields("cu04")
                  End If
               End With
               If strExc(9) <> "" Then
                  .Selection.TypeText strExc(8) & String(4, "　") & "　　　" & strExc(9)
                  .Selection.TypeParagraph
               End If
            End If
         End If
      Next
      If pa(1) = "FCT" Then
         strExc(9) = ""
         If pa(7) <> "" Then
            strExc(9) = pa(7)
         ElseIf pa(6) <> "" Then
            strExc(9) = pa(6)
         Else
            strExc(9) = pa(5)
         End If
         'Modified by Morgan 2022/7/27
         '.Selection.TypeText "　　　商標名�耤G" & strExc(9)
         '.Selection.TypeParagraph
         '.Selection.TypeText "　　　商品�P分：第" & pa(9) & "類"
         '.Selection.TypeParagraph
         '.Selection.TypeText "　　　貴方整理番�A：" & pa(45)
         '.Selection.TypeParagraph
         '.Selection.TypeText "　　　�U方整理番�A：" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4))
         strUniText = PUB_GetUniText(Me.Name, "名稱")
         .Selection.TypeText "　　　商標" & strUniText & strExc(9)
         .Selection.TypeParagraph
         strUniText = PUB_GetUniText(Me.Name, "區分")
         .Selection.TypeText "　　　商品" & strUniText & "第" & pa(9) & "類"
         .Selection.TypeParagraph
         strUniText = PUB_GetUniText(Me.Name, "彼所案號")
         .Selection.TypeText "　　　" & strUniText & pa(45)
         .Selection.TypeParagraph
         strUniText = PUB_GetUniText(Me.Name, "本所案號")
         .Selection.TypeText "　　　" & strUniText & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4))
         'end 2022/7/27
      Else
         'Add by Morgan 2010/9/7
         If Left(Text1, 1) = "C" Then
            'Modified by Morgan 2022/7/27
            '.Selection.TypeText "　　　貴方整理番�A：" & m_strCP45
            strUniText = PUB_GetUniText(Me.Name, "彼所案號")
            .Selection.TypeText "　　　" & strUniText & m_strCP45
            'end 2022/7/27
         Else
         'end 2010/9/7
            'Modified by Morgan 2022/7/27
            '.Selection.TypeText "　　　貴方整理番�A：" & pa(77)
            strUniText = PUB_GetUniText(Me.Name, "彼所案號")
            .Selection.TypeText "　　　" & strUniText & pa(77)
            'end 2022/7/27
         End If
         .Selection.TypeParagraph
         'Modified by Morgan 2022/7/27
         '.Selection.TypeText "　　　�U方整理番�A：" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4))
         strUniText = PUB_GetUniText(Me.Name, "本所案號")
         .Selection.TypeText "　　　" & strUniText & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4))
         'end 2022/7/27
      End If
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      'Modified by Morgan 2022/7/27
      '.Selection.TypeText "���K  時下�dぼ�蝎M�Xソ�磈Oシ�敯yヂ申�篟噮暷e�魽C"
      strUniText = PUB_GetUniText(Me.Name, "拜啟")
      .Selection.TypeText strUniText
      'end 2022/7/27
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeText "                                                      敬具"
      
      '2010/9/10 CFP要改為不管日文或英文或數字都用MS Mincho
      If pa(1) = "CFP" Then
         .Selection.WholeStory
         .Selection.Font.Name = "MS Mincho"
      '2010/9/10 END
      Else
         '英文字型設Times New Roman
         .Selection.WholeStory
         'Modified by Morgan 2014/9/24 改與定稿一致
         '.Selection.Font.Name = "Times New Roman"
         'Modified
         'Modified by Morgan 2022/7/13 改 MS Mincho -- 毓芳, May
         '.Selection.Font.Name = "細明體"
         .Selection.Font.Name = "MS Mincho"
         'end 2022/7/13
         'end 2014/9/24
      End If   '2010/9/10 ADD
      ChgWordFormat g_WordAp, .Selection.Text 'Added by Morgan 2015/8/18
      
   End With
    
   g_WordAp.Visible = True
   g_WordAp.WindowState = wdWindowStateMaximize
   Set g_WordAp = Nothing 'Added by Morgan 2015/9/7
   
ERRORSECTION1:
   If Err.NUMBER <> 0 Then
      Select Case Err.NUMBER
         Case 91, 462:
            If bolRetry = True Then
               MsgBox Err.Description, vbCritical
            Else
               Set g_WordAp = New Word.Application
               g_WordAp.Documents.add
               bolRetry = True
               Resume
            End If
         Case Else:
            MsgBox Err.Description, vbCritical
      End Select
   End If
End Sub

Private Function SplitTitle(stTitle As String, iReserve As Integer, Optional iMax As Integer = 82) As String
   Dim arrWords '單字陣列
   Dim ii As Integer
   Dim iUsable As Integer
   Dim iRest As Integer
   Dim strTmp As String
   Dim strWord As String
   Dim iLen As Integer
      
   iUsable = iMax - iReserve
   stTitle = Trim(stTitle)
   If stTitle <> "" Then
      arrWords = Split(stTitle, " ")
      strTmp = ""
      iRest = iUsable
       For ii = LBound(arrWords) To UBound(arrWords)
         strWord = arrWords(ii)
         Do While strWord <> ""
            iLen = PUB_GetLen(strWord)
            '超過最大可印長度時,接續印並斷字後跳行
            If iLen > iUsable Then
               strTmp = strTmp & " " & GetWord(strWord, iRest, strWord) & vbCrLf & String(iReserve, " ")
               iRest = iUsable
            '超過剩餘最大可印長度時,跳行後列印
            ElseIf iLen > iRest - 1 Then
               strTmp = strTmp & vbCrLf & String(iReserve, " ") & strWord
               iRest = iUsable - iLen
               strWord = ""
            '可印
            Else
               If iRest = iUsable Then
                  strTmp = strTmp & strWord
                  iRest = iRest - iLen
               Else
                  strTmp = strTmp & " " & strWord
                  iRest = iRest - iLen - 1
               End If
               strWord = ""
            End If
         Loop
       Next
   End If
   SplitTitle = strTmp
End Function

'Modify By Sindy 2024/7/23 mark,改用共用函數 PUB_GetLen
'Private Function GetLen(strWord As String) As Integer
'   Dim stChar As String
'   Dim iLen As Integer
'   Dim ii As Integer
'   For ii = 1 To Len(strWord)
'      stChar = Mid(strWord, ii, 1)
'      '全形字 2
'      If Asc(stChar) < 0 Then
'         iLen = iLen + 2
'      '英文大寫 1.5
'      ElseIf Asc(stChar) >= 65 And Asc(stChar) <= 90 Then
'         iLen = iLen + 1.5
'      '其他 1
'      Else
'         iLen = iLen + 1
'      End If
'   Next
'   GetLen = iLen
'End Function

Private Function GetWord(ByVal strWord As String, ByVal iMaxLen As Integer, ByRef strWord2 As String) As String
   Dim stChar As String
   Dim iLen As Integer
   Dim strTmp As String
   Dim ii As Integer
   
   For ii = 1 To Len(strWord)
      stChar = Mid(strWord, ii, 1)
      '全形字 2
      If Asc(stChar) < 0 Then
         iLen = iLen + 2
      '英文大寫 1.5
      ElseIf Asc(stChar) >= 65 And Asc(stChar) <= 90 Then
         iLen = iLen + 1.5
      '其他 1
      Else
         iLen = iLen + 1
      End If
      If iLen > iMaxLen Then
         Exit For
      Else
         strTmp = strTmp & stChar
      End If
   Next
   If strTmp = strWord Then
      strWord2 = ""
   Else
      strWord2 = Mid(strWord, ii)
   End If
   GetWord = strTmp
End Function

'Modify By Sindy 2014/9/18 Move basLetter
''Add by Morgan 2008/4/8
''讀取EMail信箱
'Private Function GetEMail() As String
'   If Option2.Value = True Then
'      strExc(1) = m_strFCAgent
'   ElseIf Option6.Value = True Then
'      strExc(1) = m_strCP44
'      'Added by Morgan 2012/6/25
'      'CFP 日本案 CF代理人為 Y51835 E-Mail 用 mm@miyoshipat.co.jp --慧汶 101/5/30 請作單
'      If Text1 = "CFP" And Label11 = "011" And m_strCP44 = "Y51835000" Then
'         GetEMail = "mm@miyoshipat.co.jp"
'         Exit Function
'      End If
'   Else
'      strExc(1) = m_CustNo(1)
'   End If
'   If Left(strExc(1), 1) = "Y" Then
'      strExc(0) = "SELECT FA16 FROM FAGENT WHERE " & ChgFagent(strExc(1))
'   Else
'      strExc(0) = "SELECT CU20 FROM CUSTOMER WHERE " & ChgCustomer(strExc(1))
'   End If
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      With RsTemp
'         GetEMail = "" & .Fields(0)
'      End With
'   End If
'End Function

'Add by Morgan 2008/8/1
'取得個案聯絡人編號
Private Function GetCaseContactNo(p_CaseNo1 As String, p_CaseNo2 As String, p_CaseNo3 As String, p_CaseNo4 As String, p_CustNo As String) As String
   Dim stSysNo As String, stSQL As String, intR As Integer, stCaseNo As String
   stCaseNo = Trim(p_CaseNo1 & p_CaseNo2 & p_CaseNo3 & p_CaseNo4)
   '若有本所案號
   If stCaseNo <> "" And p_CustNo <> "" Then
     stSysNo = CheckSys(p_CaseNo1)
     Select Case stSysNo
        Case "1"
            stSQL = "select pa149 from patent where " & ChgPatent(stCaseNo) & " and pa26='" & ChangeCustomerL(p_CustNo) & "'"
        Case "2"
            stSQL = "select tm123 from trademark where " & ChgTradeMark(stCaseNo) & " and tm23='" & ChangeCustomerL(p_CustNo) & "'"
        Case "3"
            stSQL = "select lc42 from Lawcase where " & ChgLawcase(stCaseNo) & " and lc11='" & ChangeCustomerL(p_CustNo) & "'"
        Case "4"
            stSQL = "select hc23 from Hirecase where " & ChgHirecase(stCaseNo) & " and hc05='" & ChangeCustomerL(p_CustNo) & "'"
        Case Else
            stSQL = "select sp78 from Servicepractice where " & ChgService(stCaseNo) & " and sp08='" & ChangeCustomerL(p_CustNo) & "'"
     End Select
     intR = 1
     Set RsTemp = ClsLawReadRstMsg(intR, stSQL)
     If intR = 1 Then
        GetCaseContactNo = "" & RsTemp(0)
     End If
   End If
End Function

'調整首行凸排
Sub PhaseIndent()
    g_WordAp.Selection.WholeStory
    With g_WordAp.Selection.ParagraphFormat
        .LeftIndent = g_WordAp.CentimetersToPoints(1)
        .RightIndent = g_WordAp.CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 15
        '2015/7/15 cancel by sonia 取消分散對齊,否則落款的右靠無效
        '.Alignment = wdAlignParagraphJustify
        .WidowControl = False
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = g_WordAp.CentimetersToPoints(-1)
        .OutlineLevel = wdOutlineLevelBodyText
        .AutoAdjustRightIndent = True
        .DisableLineHeightGrid = True
        .FarEastLineBreakControl = True
        .WordWrap = True
        .HangingPunctuation = True
        .HalfWidthPunctuationOnTopOfLine = False
        .AddSpaceBetweenFarEastAndAlpha = True
        .AddSpaceBetweenFarEastAndDigit = True
        .BaseLineAlignment = wdBaselineAlignAuto
    End With
End Sub

'Add by Morgan 2007/2/7
'讀取信函備註
'Remove by Lydia 2019/09/27 改成共用->basPublc.Pub_GetLetterMemo
'Private Function GetLetterMemo(p_Sys As String, p_Language As String) As String
'   Dim stSQL As String, intR As Integer
'   stSQL = "select LM05 from lettermemo where LM01='" & pa(1) & "' and LM02='" & p_Language & "' and (" & strSrvDate(1) & " between LM03 and LM04 )"
'   intR = 1
'   Set RsTemp = ClsLawReadRstMsg(intR, stSQL)
'   If intR = 1 Then
'      GetLetterMemo = "" & RsTemp.Fields(0)
'   End If
'End Function

'Added by Morgan 2018/2/7
Private Sub InsQA()
   Dim oTable As Word.Table
   Dim oShape1 As Word.Shape, oShape2 As Word.Shape, oShape3 As Word.Shape, oShape4 As Word.Shape
   Dim pageX As Double '頁面X軸位置
   Dim pageY As Double '頁面Y軸位置
   
   With g_WordAp
   .Visible = True
   .Selection.TypeParagraph
   Set oTable = .ActiveDocument.Tables.add(Range:=.Selection.Range, NumRows:=2, NumColumns:=1)
   End With
   With oTable
   .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
   .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
   .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
   .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
   .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
   .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
   .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
   .Borders.Shadow = False
   '.Shading.BackgroundPatternColorIndex = RGB(239, 223, 236)
   .Shading.BackgroundPatternColorIndex = wdTurquoise
   End With
   oTable.Select
   With g_WordAp.Selection
   .Collapse Direction:=wdCollapseStart
   .Font.Name = "Times New Roman"
   .Font.Italic = True
   .Font.Size = 12
   .TypeText Text:="Note: In order to further enhance the quality of our service, the Quality Assurance Department of Tai E invites you to participate in our Service Quality Survey by two steps: (1) "
   .Font.Bold = True
   .TypeText Text:="Clicking "
   .Font.Bold = False
   .TypeText Text:="one of the following icons; (2) "
   .Font.Bold = True
   .TypeText Text:="Sending "
   .Font.Bold = False
   .TypeText Text:="the pop-out email. If you have any additional comments, please write them in the body of the pop-out email before sending. Thank you for your participation in advance."
   .MoveDown Unit:=wdLine, Count:=1
   .Cells(1).SetHeight RowHeight:=30, HeightRule:=wdRowHeightAtLeast
   pageX = .Information(wdHorizontalPositionRelativeToPage)
   pageY = .Information(wdVerticalPositionRelativeToPage)
   End With
   
   Set oShape1 = g_WordAp.ActiveDocument.Shapes.AddShape(5, pageX, pageY, 57, 20)
   oShape1.Left = g_WordAp.CentimetersToPoints(3)
   oShape1.Top = g_WordAp.CentimetersToPoints(0.13)
   
   Set oShape2 = g_WordAp.ActiveDocument.Shapes.AddShape(5, pageX, pageY, 57, 20)
   oShape2.Left = oShape1.Left + oShape1.Width + g_WordAp.CentimetersToPoints(0.2)
   oShape2.Top = oShape1.Top
   
   Set oShape3 = g_WordAp.ActiveDocument.Shapes.AddShape(5, pageX, pageY, 57, 20)
   oShape3.Left = oShape2.Left + oShape2.Width + g_WordAp.CentimetersToPoints(0.2)
   oShape3.Top = oShape1.Top
   
   Set oShape4 = g_WordAp.ActiveDocument.Shapes.AddShape(5, pageX, pageY, 57, 20)
   oShape4.Left = oShape3.Left + oShape3.Width + g_WordAp.CentimetersToPoints(0.2)
   oShape4.Top = oShape1.Top
   
   'Excellent
   oShape1.Select
   With g_WordAp.Selection
   .ShapeRange.TextFrame.MarginLeft = 0#
   .ShapeRange.TextFrame.MarginRight = 0#
   .ShapeRange.TextFrame.MarginTop = 0#
   .ShapeRange.TextFrame.MarginBottom = 0#
   .ShapeRange.Fill.ForeColor.RGB = RGB(255, 0, 0)
   End With
   oShape1.TextFrame.TextRange.Select
   With g_WordAp.Selection
   .ParagraphFormat.Alignment = wdAlignParagraphCenter
   .Font.Name = "Arial"
   .Font.Size = 10
   .Font.Bold = True
   .Font.Underline = wdUnderlineSingle
   .Font.ColorIndex = wdBlue
   .TypeText Text:="Excellent"
   .Hyperlinks.add Anchor:=.ShapeRange, address:= _
        "mailto:qadept@taie.com.tw?subject=[QoS]" & strUserNum & ";%20" & CaseNo & ";%20Excellent", SubAddress:=""
   End With
   
   'Good
   oShape2.Select
   With g_WordAp.Selection
   .ShapeRange.TextFrame.MarginLeft = 0#
   .ShapeRange.TextFrame.MarginRight = 0#
   .ShapeRange.TextFrame.MarginTop = 0#
   .ShapeRange.TextFrame.MarginBottom = 0#
   .ShapeRange.Fill.ForeColor.RGB = RGB(255, 255, 0)
   End With
   oShape2.TextFrame.TextRange.Select
   With g_WordAp.Selection
   .ParagraphFormat.Alignment = wdAlignParagraphCenter
   .Font.Name = "Arial"
   .Font.Size = 10
   .Font.Bold = True
   .Font.Underline = wdUnderlineSingle
   .Font.ColorIndex = wdBlue
   .TypeText Text:="Good"
   .Hyperlinks.add Anchor:=.ShapeRange, address:= _
        "mailto:qadept@taie.com.tw?subject=[QoS]" & strUserNum & ";%20" & CaseNo & ";%20Good", SubAddress:=""
   End With
   
   'Fair
   oShape3.Select
   With g_WordAp.Selection
   .ShapeRange.TextFrame.MarginLeft = 0#
   .ShapeRange.TextFrame.MarginRight = 0#
   .ShapeRange.TextFrame.MarginTop = 0#
   .ShapeRange.TextFrame.MarginBottom = 0#
   .ShapeRange.Fill.ForeColor.RGB = RGB(146, 208, 80)
   End With
   
   oShape3.TextFrame.TextRange.Select
   With g_WordAp.Selection
   .ParagraphFormat.Alignment = wdAlignParagraphCenter
   .Font.Name = "Arial"
   .Font.Size = 10
   .Font.Bold = True
   .Font.Underline = wdUnderlineSingle
   .Font.ColorIndex = wdBlue
   .TypeText Text:="Fair"
   .Hyperlinks.add Anchor:=.ShapeRange, address:= _
        "mailto:qadept@taie.com.tw?subject=[QoS]" & strUserNum & ";%20" & CaseNo & ";%20Fair", SubAddress:=""
   End With
   
   'Poor
   oShape4.Select
   With g_WordAp.Selection
   .ShapeRange.TextFrame.MarginLeft = 0#
   .ShapeRange.TextFrame.MarginRight = 0#
   .ShapeRange.TextFrame.MarginTop = 0#
   .ShapeRange.TextFrame.MarginBottom = 0#
   .ShapeRange.Fill.ForeColor.RGB = RGB(0, 176, 240)
   End With
   
   oShape4.TextFrame.TextRange.Select
   With g_WordAp.Selection
   .ParagraphFormat.Alignment = wdAlignParagraphCenter
   .Font.Name = "Arial"
   .Font.Size = 10
   .Font.Bold = True
   .Font.Underline = wdUnderlineSingle
   .Font.ColorIndex = wdBlue
   .TypeText Text:="Poor"
   .Hyperlinks.add Anchor:=.ShapeRange, address:= _
        "mailto:qadept@taie.com.tw?subject=[QoS]" & strUserNum & ";%20" & CaseNo & ";%20Poor", SubAddress:=""
   End With
   g_WordAp.Selection.EndKey Unit:=wdStory
End Sub

'Added by Morgan 2011/11/18
'外專人員加印代理人備註及案件備註
Private Sub PrintMemo()
Dim intQ As Integer 'Added by Lydia 2021/02/08

   '外專人員要帶出案件備註及代理人備註
   'Modified by Morgan 2012/1/6 調整順序改先案件備註後代理人備註(原來相反)--黃美珍
   If Left(m_StrUserST03, 2) = "F2" Then
   'If (Left(m_StrUserST03, 2) = "F2" And strSrvDate(1) < 各項指示啟用日) Or strSrvDate(1) >= 各項指示啟用日 Then 'Mark by Lydia 2020/08/27 保留（各項指示：改成不限使用者部門）
      'Added by Morgan 2018/2/7 外專工程師+滿意度調查
      If m_StrUserST03 = "F21" Then
         'InsQA '等 Elvan 通知再上線
      End If
      'end 2018/2/7
   
      'Memo by Lydia 2020/08/26 各項指示：尚未完成確認，先帶出原先備註後面再帶出各項指示
      If strSrvDate(1) < 各項指示啟用日 Or (strSrvDate(1) >= 各項指示啟用日 And PUB_GetInstConfirm(m_StrUserST03, pa(1) & pa(2) & pa(3) & pa(4))) = False Then 'Added by Lydia 2020/08/26 增加判斷
          'Modified by Lydia 2020/12/16 若無案件備註不用顯示；ex.FCP-62881的個案備註已清空，但是各項指示未做完成確認
          'If Text5 <> "" Then
          strExc(1) = IIf(strKeyCase <> "", Replace(Text5, strMemoCase, ""), Text5)
          If Trim(strExc(1)) <> "" Then
          'end 2020/12/16
             g_WordAp.Selection.TypeParagraph
             g_WordAp.Selection.Font.ColorIndex = wdRed
             g_WordAp.Selection.TypeText "案件備註："
             g_WordAp.Selection.TypeParagraph 'Added by Lydia 2020/07/27
             g_WordAp.Selection.Font.ColorIndex = wdDarkBlue
             'Modified by Lydia 2020/06/04 拿掉各項指示
             'g_WordAp.Selection.TypeText Text5
             g_WordAp.Selection.TypeText IIf(strKeyCase <> "", Replace(Text5, vbCrLf & strMemoCase, ""), Text5)
             g_WordAp.Selection.Font.ColorIndex = wdAuto
          End If
      End If 'Added by Lydia 2020/08/26 增加判斷
      'Added by Lydia 2020/06/04 備註和各項指示並存(暫時到完全取代)
      'Modified by Lydia 2021/02/08 增加各項指示勾選項判斷
      If strKeyCase <> "" And (ChkINST.Value = 1 Or ChkINST.Visible = False) Then
         PrintMemoTable "3", strKeyCase
      End If
      
      'Memo by Lydia 2020/08/26 各項指示：尚未完成確認，先帶出原先備註後面再帶出各項指示
      If strSrvDate(1) < 各項指示啟用日 Or (strSrvDate(1) >= 各項指示啟用日 And PUB_GetInstConfirm(m_StrUserST03, IIf(Option6.Value = True, m_strCP44, m_strFCAgent))) = False Then 'Added by Lydia 2020/08/26 增加判斷
        'Modified by Lydia 2020/12/16 若無案件備註不用顯示；
        'If Text6(0) <> "" Then
        strExc(1) = IIf(strKeyY <> "", Replace(Text6(0), strMemoY, ""), Text6(0))
        If Trim(strExc(1)) <> "" Then
        'end 2020/12/16
           g_WordAp.Selection.TypeParagraph
           g_WordAp.Selection.Font.ColorIndex = wdRed
           g_WordAp.Selection.TypeText "代理人備註："
           g_WordAp.Selection.TypeParagraph 'Added by Lydia 2020/07/27
           g_WordAp.Selection.Font.ColorIndex = wdDarkBlue
           'Modified by Lydia 2020/06/04 拿掉各項指示
           'g_WordAp.Selection.TypeText Text6(0)
           g_WordAp.Selection.TypeText IIf(strKeyY <> "", Replace(Text6(0), vbCrLf & strMemoY, ""), Text6(0))
           g_WordAp.Selection.Font.ColorIndex = wdAuto
        End If
      End If 'Added by Lydia 2020/08/26 增加判斷
      'Added by Lydia 2020/06/04 備註和各項指示並存(暫時到完全取代)
      'Modified by Lydia 2021/02/08 增加各項指示勾選項判斷
      If strKeyY <> "" And (ChkINST.Value = 1 Or ChkINST.Visible = False) Then
         PrintMemoTable "1", strKeyY
      End If

      'Memo by Lydia 2020/08/26 各項指示：尚未完成確認，先帶出原先備註後面再帶出各項指示
      If strSrvDate(1) < 各項指示啟用日 Or (strSrvDate(1) >= 各項指示啟用日 And PUB_GetInstConfirm(m_StrUserST03, m_CustNo(1))) = False Then 'Added by Lydia 2020/08/26 增加判斷
        'Added by Morgan 2014/7/31
        'Modified by Lydia 2020/12/16 若無案件備註不用顯示；
        'If Text6(1) <> "" Then
        strExc(1) = IIf(strKeyX <> "", Replace(Text6(1), strMemoX, ""), Text6(1))
        If Trim(strExc(1)) <> "" Then
        'end 2020/12/16
           g_WordAp.Selection.TypeParagraph
           g_WordAp.Selection.Font.ColorIndex = wdRed
           'Modified by Lydia 2021/02/08 改抬頭
           'g_WordAp.Selection.TypeText "申請人備註："
           g_WordAp.Selection.TypeText "申請人1（" & m_CustName(1) & "）備註："
           g_WordAp.Selection.TypeParagraph 'Added by Lydia 2020/07/27
           g_WordAp.Selection.Font.ColorIndex = wdDarkBlue
           'Modified by Lydia 2020/06/04 拿掉各項指示
           'g_WordAp.Selection.TypeText Text6(1)
           g_WordAp.Selection.TypeText IIf(strKeyX <> "", Replace(Text6(1), vbCrLf & strMemoX, ""), Text6(1))
           g_WordAp.Selection.Font.ColorIndex = wdAuto
        End If
        'end 2014/7/31
      End If 'Added by Lydia 2020/08/26
      'Added by Lydia 2020/06/04 備註和各項指示並存(暫時到完全取代)
      'Modified by Lydia 2021/02/08 增加各項指示勾選項判斷
      If strKeyX <> "" And (ChkINST.Value = 1 Or ChkINST.Visible = False) Then
         PrintMemoTable "2", strKeyX
      End If
      
      'Added by Lydia 2021/02/08 針對國外部原本只顯示申請人1的備註+各項指示，現在顯示申請人1~5的備註+各項指示
      If ChkINST.Value = 1 Then
         For intQ = 2 To 5
            If Trim(m_CustNo(intQ)) <> "" Then
                g_WordAp.Selection.TypeParagraph
                If m_CustMemo(intQ) <> "" And PUB_GetInstConfirm(m_StrUserST03, m_CustNo(intQ)) = False Then '有備註，尚未有「完成確認」
                    g_WordAp.Selection.Font.ColorIndex = wdRed
                    g_WordAp.Selection.TypeText "申請人" & intQ & "（" & m_CustName(intQ) & "）備註："
                    g_WordAp.Selection.TypeParagraph
                    g_WordAp.Selection.Font.ColorIndex = wdDarkBlue
                    g_WordAp.Selection.TypeText m_CustMemo(intQ)
                    g_WordAp.Selection.Font.ColorIndex = wdAuto
                End If
                PrintMemoTable "2", m_CustNo(intQ), intQ
            End If
         Next intQ
      End If
      'end 2021/02/08
      
      If Text6(0) & Text5 & Text6(1) <> "" Then
         g_WordAp.Selection.TypeParagraph
      End If

   End If
End Sub

'Added by Lydia 2017/08/03 各項指示(David-頁籤)
'Modified by Lydia 2021/02/08 +申請人1~5 (iSeq)
Private Sub PrintMemoTable(ByVal iKind As String, ByVal iKeyNo As String, Optional ByVal iSeq As Integer = "1")
Dim intJ As Integer, intK As Integer
Dim arrTmp As Variant, arrVarTmp As Variant
Dim strMid As String
Dim strGrp As String

On Error Resume Next 'Added by Lydia 2020/08/31 中文有時會出現不知名錯誤,沒有出Word並且重新再開Word檔
     
    'Modified by Lydia 2021/02/08 判斷是否含欄位設定=基本檔
    'If Pub_GetInstructions(Me.Name, iKeyNo, strMid, , "W", , m_strIT10) = True Then
    If Pub_GetInstructions(Me.Name, iKeyNo, strMid, , "W", , m_strIT10, IIf(ChkINSTdef.Value = 1, "0", "1")) = False Then
        g_WordAp.Selection.TypeParagraph  '沒備註,加空白列
    Else
    'end 2021/02/08
        'Move by Lydia 2021/02/08 從外面移進來
        g_WordAp.Selection.TypeParagraph
        g_WordAp.Selection.Font.ColorIndex = wdRed
        'Modified by Lydia 2020/06/04
        'g_WordAp.Selection.TypeText IIf(iKind = "1", "代理人備註：", IIf(iKind = "2", "申請人備註：", "案件備註："))
        'Memo by Lydia 2021/02/08 改抬頭
        'g_WordAp.Selection.TypeText IIf(iKind = "1", "代理人各項指示：", IIf(iKind = "2", "申請人各項指示：", "案件各項指示："))
        If iKind = "1" Then
            g_WordAp.Selection.TypeText "代理人各項指示："
        ElseIf iKind = "2" Then
            g_WordAp.Selection.TypeText "申請人" & iSeq & "（" & m_CustName(iSeq) & "）各項指示："
        ElseIf iKind = "3" Then
            g_WordAp.Selection.TypeText "案件各項指示："
        End If
        
        g_WordAp.Selection.Font.ColorIndex = wdAuto
        g_WordAp.Selection.TypeParagraph
        'end 2021/02/08 '----'Move by Lydia 2021/02/08 從外面移進來
        arrTmp = Empty
        arrTmp = Split(strMid, "|;|") '分隔記錄|;|
        '新增表格(1X3)
        g_WordAp.Selection.Tables.add Range:=g_WordAp.Selection.Range, NumRows:=1, NumColumns:=3
        g_WordAp.Selection.SelectRow
        g_WordAp.Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
        g_WordAp.Selection.Borders(wdBorderRight).LineStyle = wdLineStyleSingle
        g_WordAp.Selection.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
        g_WordAp.Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
        g_WordAp.Selection.Borders(wdBorderHorizontal).LineStyle = wdLineStyleSingle
        g_WordAp.Selection.Borders(wdBorderVertical).LineStyle = wdLineStyleSingle
        g_WordAp.Selection.Borders.Shadow = False
        
        g_WordAp.Selection.SelectRow
        g_WordAp.Selection.Cells.VerticalAlignment = wdAlignVerticalTop
        g_WordAp.Selection.Paragraphs.Alignment = wdAlignParagraphLeft
        'Added by Lydia 2020/08/26
        If Option1(0).Value = True Then '中文的左右邊界2公分 , 標楷體=14
            g_WordAp.Selection.Cells(1).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(3), RulerStyle:=wdAdjustProportional
            g_WordAp.Selection.Cells(2).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(11.5), RulerStyle:=wdAdjustProportional
            g_WordAp.Selection.Cells(3).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(2.5), RulerStyle:=wdAdjustProportional
        ElseIf Option1(2).Value = True Then '日文的左右邊界3.17公分,細明體=12, 較Times New Roman大
            g_WordAp.Selection.Cells(1).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(2.5), RulerStyle:=wdAdjustProportional
            g_WordAp.Selection.Cells(2).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(11), RulerStyle:=wdAdjustProportional
            g_WordAp.Selection.Cells(3).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(2.2), RulerStyle:=wdAdjustProportional
        Else   '英文的左右邊界3.17公分,Times New Roman=12
        'end 2020/08/26
            g_WordAp.Selection.Cells(1).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(2.5), RulerStyle:=wdAdjustProportional
            g_WordAp.Selection.Cells(2).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(11), RulerStyle:=wdAdjustProportional
            g_WordAp.Selection.Cells(3).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
        End If 'end 2020/08/26
        
        For intJ = 0 To UBound(arrTmp)
            If Trim(arrTmp(intJ)) <> "" Then
                arrVarTmp = Empty
                arrVarTmp = Split(Trim(arrTmp(intJ)), "|+|") '分隔欄位|+|
                g_WordAp.Selection.Collapse Direction:=wdCollapseStart
                For intK = 0 To UBound(arrVarTmp)
                    If Trim(arrVarTmp(intK)) <> "" Then
                        Select Case intK
                            Case 0
                                'Modified by Lydia 2020/06/11 預設格式: 文字靠左,首行凸排0.3公分
                                'g_WordAp.Selection.Paragraphs.Alignment = wdAlignParagraphCenter
                                If intJ = 0 Then
                                    g_WordAp.Selection.Paragraphs.Alignment = wdAlignParagraphLeft
                                    g_WordAp.Selection.ParagraphFormat.LeftIndent = CentimetersToPoints(0.3)
                                    g_WordAp.Selection.ParagraphFormat.RightIndent = CentimetersToPoints(0)
                                    g_WordAp.Selection.ParagraphFormat.FirstLineIndent = CentimetersToPoints(-0.3)
                                End If
                                'end 2020/06/11
                                If strGrp <> arrVarTmp(intK) Then
                                   g_WordAp.Selection.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
                                   g_WordAp.Selection.TypeText Text:=Trim(arrVarTmp(intK))
                                Else
                                   g_WordAp.Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
                                End If
                                strGrp = arrVarTmp(intK)
                            Case 1, 2
                                'Added by Lydia 2020/06/11 預設格式:文字置中
                                If intJ = 0 And intK = 2 Then
                                    g_WordAp.Selection.Paragraphs.Alignment = wdAlignParagraphCenter
                                End If
                                'end 2020/06/11
                                g_WordAp.Selection.TypeText Text:=Trim(arrVarTmp(intK))
                        End Select
                    End If
                    g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
                Next intK
                
                If intJ < UBound(arrTmp) Then
                   g_WordAp.Selection.InsertRows
                   g_WordAp.Selection.Collapse Direction:=wdCollapseStart
                End If
            End If
        Next intJ
        g_WordAp.Selection.EndKey Unit:=wdStory
    End If
    
End Sub

'Modify By Sindy 2014/9/18 Move basQuery
''Added by Morgan 2013/3/5
''國家英文名(要和 basLetter 同步修改)
'Private Function GetNationEngName(pCode As String) As String
'   Dim stSQL As String, intQ As Integer
'   Dim rsQuery As ADODB.Recordset
'
'   stSQL = "select decode(na01,'000','Taiwan','020','PRC',decode(instr('101,201',na01),0,null,'the ')||initcap(NA04)) from nation where na01='" & pCode & "'"
'   intQ = 1
'   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
'   If intQ = 1 Then
'      GetNationEngName = "" & rsQuery(0)
'   End If
'   Set rsQuery = Nothing
'End Function

'ADD BY SONIA 2015/9/16 依條款代碼取得條款名稱caselaw(自Command1_Click之核駁前先行通知1202抽出來)
Private Sub GetLaw()
Dim caselawcount As String    '來函條款個數 　2008/11/19 ADD BY SONIA
Dim caselawTemp As Variant                   '2008/11/19 ADD BY SONIA
Dim strLawItem As String
                  
   caselaw = ""
   caselawTemp = Split(m_CP49, ",")
   caselawcount = UBound(caselawTemp) + 1
   strLawItem = ""
   For i = 0 To UBound(caselawTemp)
      If strLawItem <> "" Then strLawItem = strLawItem & ","
      If Len(caselawTemp(i)) = 1 Then
         strLawItem = strLawItem & caselawTemp(i)
      Else
         strLawItem = strLawItem & Mid(caselawTemp(i), 2, 1)
      End If
   Next i
   If caselawcount = 3 Then
      '第23條第1項第1款(AXX)及第23條第1項第2款(BXX)及第23條第1項第11款(FXX)
      If InStr(strLawItem, "A") > 0 And InStr(strLawItem, "B") > 0 And InStr(strLawItem, "F") > 0 Then
         If m_Combo8 = "00" Then m_Combo8 = "2C"
         caselaw = "第5條第2項及第23條第1項第1款、第2款及第11款"
      '第23條第1項第1款(AXX)及第23條第1項第2款(BXX)及第23條第1項第13款(HXX)
      ElseIf InStr(strLawItem, "A") > 0 And InStr(strLawItem, "B") > 0 And InStr(strLawItem, "H") > 0 Then
         If m_Combo8 = "00" Then m_Combo8 = "2D"
         caselaw = "第5條第2項及第23條第1項第1款、第2款及第13款"
      End If
   ElseIf caselawcount = 2 Then
      'Modify By Sindy 2012/9/6
      '第29條第1項第1款(AXX)及第29條第1項第3款(BXX)
      If InStr(strLawItem, "A") > 0 And InStr(strLawItem, "B") > 0 Then
         If m_Combo8 = "00" Then m_Combo8 = "21"
         caselaw = "第18條第2項及第29條第1項第1款及第3款"
      '2012/9/6 End
      'Add By Sindy 2012/9/6
      '第29條第1項第2款(AXX)及第29條第1項第3款(CXX)
      ElseIf InStr(strLawItem, "A") > 0 And InStr(strLawItem, "C") > 0 Then
         If m_Combo8 = "00" Then m_Combo8 = "2M"
         caselaw = "第18條第2項、第29條第1項第2款及第3款"
      '2012/9/6 End
      '第29條第1項第3款(AXX)及第30條第1項第10款(HXX)
      ElseIf InStr(strLawItem, "A") > 0 And InStr(strLawItem, "H") > 0 Then
         If m_Combo8 = "00" Then m_Combo8 = "2E"
         caselaw = "第18條第2項、第29條第1項第2款及第30條第1項第10款"
      '第29條第1項第1款(BXX)及第30條第1項第8款(FXX)
      ElseIf InStr(strLawItem, "B") > 0 And InStr(strLawItem, "F") > 0 Then
         If m_Combo8 = "00" Then m_Combo8 = "22"
         caselaw = "第18條第2項、第29條第1項第1款及第30條第1項第8款"
      '第29條第1項第1款(BXX)及第30條第1項第10款(HXX)
      ElseIf InStr(strLawItem, "B") > 0 And InStr(strLawItem, "H") > 0 Then
         If m_Combo8 = "00" Then m_Combo8 = "23"
         caselaw = "第18條第2項、第29條第1項第1款及第30條第1項第10款"
      '第30條第1項第8款(FXX)及第30條第1項第10款(HXX)
      ElseIf InStr(strLawItem, "F") > 0 And InStr(strLawItem, "H") > 0 Then
         If m_Combo8 = "00" Then m_Combo8 = "24"
         caselaw = "第30條第1項第8款及第10款"
      '第29條第1項第2款(CXX)及第30條第1項第8款(FXX)
      ElseIf InStr(strLawItem, "C") > 0 And InStr(strLawItem, "F") > 0 Then
         If m_Combo8 = "00" Then m_Combo8 = "2G"
         caselaw = "第18條第2項、第29條第1項第2款及第30條第1項第8款"
      '第29條第1項第3款(AXX)及第30條第1項第8款(FXX)
      ElseIf InStr(strLawItem, "A") > 0 And InStr(strLawItem, "F") > 0 Then
         If m_Combo8 = "00" Then m_Combo8 = "2H"
         caselaw = "第18條第2項、第29條第1項第3款及第30條第1項第8款"
      '第29條第3項(MXX)及第30條第1項第8款(FXX)
      ElseIf InStr(strLawItem, "M") > 0 And InStr(strLawItem, "F") > 0 Then
         If m_Combo8 = "00" Then m_Combo8 = "2J"
         caselaw = "第18條第2項、第29條第3項及第30條第1項第8款"
      '第29條第3項(MXX)及第30條第1項第10款(HXX)
      ElseIf InStr(strLawItem, "M") > 0 And InStr(strLawItem, "H") > 0 Then
         If m_Combo8 = "00" Then m_Combo8 = "2K"
         caselaw = "第18條第2項、第29條第3項及第30條第1項第10款"
      'Add By Sindy 2012/7/17
      '第29條第3項(MXX)及第30條第1項第11款(GXX)
      ElseIf InStr(strLawItem, "M") > 0 And InStr(strLawItem, "G") > 0 Then
         If m_Combo8 = "00" Then m_Combo8 = "2L"
         caselaw = "第18條第2項、第29條第3項及第30條第1項11款"
      '2012/7/17 End
      End If
   ElseIf caselawcount = 1 Then
      '第29條第1項第3款(AXX)
      If InStr(strLawItem, "A") > 0 Then
         If m_Combo8 = "00" Then m_Combo8 = "25"
         caselaw = "第18條第2項及第29條第1項第3款"
      '第29條第1項第1款(BXX)
      ElseIf InStr(strLawItem, "B") > 0 Then
         If m_Combo8 = "00" Then m_Combo8 = "26"
         caselaw = "第18條第2項及第29條第1項第1款"
      '第29條第1項第2款(CXX)
      ElseIf InStr(strLawItem, "C") > 0 Then
         If m_Combo8 = "00" Then m_Combo8 = "2F"
         caselaw = "第18條第2項及第29條第1項第2款"
      '第30條第1項第8款(FXX)
      ElseIf InStr(strLawItem, "F") > 0 Then
         If m_Combo8 = "00" Then m_Combo8 = "27"
         caselaw = "第30條第1項第8款"
      '第30條第1項第11款(GXX)
      ElseIf InStr(strLawItem, "G") > 0 Then
         If m_Combo8 = "00" Then m_Combo8 = "28"
         caselaw = "第30條第1項11款"
      '第30條第1項第10款(HXX)
      ElseIf InStr(strLawItem, "H") > 0 Then
         If m_Combo8 = "00" Then m_Combo8 = "29"
         caselaw = "第30條第1項第10款"
      '第30條第1項第12款(IXX)
      ElseIf InStr(strLawItem, "I") > 0 Then
         If m_Combo8 = "00" Then m_Combo8 = "2A"
         caselaw = "第30條第1項第12款"
      '第30條第1項第13款(JXX)
      ElseIf InStr(strLawItem, "J") > 0 Then
         If m_Combo8 = "00" Then m_Combo8 = "2B"
         caselaw = "第30條第1項第13款"
      '第29條第3項(MXX)
      ElseIf InStr(strLawItem, "M") > 0 Then
         If m_Combo8 = "00" Then m_Combo8 = "2I"
         caselaw = "第18條第2項及第29條第3項"
      End If
   End If
   If m_Combo8 = "00" Then
      '其他條款(AXX)
      If m_Combo8 = "00" Then m_Combo8 = "2Z"
      caselaw = "第XX條第X項第X款"
   End If

End Sub
'END 2015/9/16

'Added by Lydia 2015/10/30 FC翻譯案件郵件
'Modified by Lydia 2015/11/11 +iCP09
'Mark by Lydia 2025/03/13 已不再使用
'Private Sub Translate_SendMail(ByVal FT14 As String, FT18 As String, strCP48 As String, iCp09 As String)
'Dim objOutLook As Object
'Dim objMail As Object
'Dim strName As String, strText As String
'Dim m_TempFileName As String
'Dim strContent As String
'Dim strPath As String
''Added by Lydia 2015/11/11
''Dim strLoadPath As String '讀取附件路徑
'Dim strFile As String
'Dim fs, f
'Dim stReName As String
''end 2015/11/11
'Dim inX As Integer 'Added by Lydia 2015/12/16
'Dim strPWD As String 'Added by Lydia 2015/12/22
'Dim rsAD As New ADODB.Recordset 'Added by Lydia 2016/07/07
'Dim strConUser As String 'Added by Lydia 2016/07/07
'
'On Error GoTo ErrHand
'
'   cmdFCMail(2).Enabled = False  'Added by Lydia 2015/12/18
'
'   strPWD = "" 'Added by Lydia 2015/12/22
'
'   'Added by Lydia 2016/07/07 抓工作清單中的台一聯絡人資料
'   'modify by sonia 2016/7/15
'   'strExc(0) = "86013"
'   strExc(0) = Pub_GetSpecMan("M")
'   strSql = "select st01,st02,st03,st04,st05,st06,st07,st22 from staff where st01='" & strExc(0) & "' "
'   intI = 0
'   Set rsAD = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      If rsAD.Fields("st04") = "2" Then
'         MsgBox "工作清單中的台一聯絡人: " & rsAD.Fields("st02") & " 已離職，請自行修改工作清單!"
'      End If
'   Else
'      Exit Sub
'   End If
'
'   'Modified by Lydia 2017/06/16 直接放在程式下面
'   'strPath = "C:\外翻工作通知單"
'   'If Dir(strPath, vbDirectory) = "" Then
'   '   MkDir strPath
'   'End If
'   'Modified by Lydia 2018/06/25 改路徑
'   'strPath = App.path
'   strPath = strSavePath1
'
'   'Added by Lydia 2015/11/11 抓本機電腦符合外翻名稱的資料夾路徑
' '  strLoadPath = Dir(strLoadPath, vbDirectory)
'
'   '判斷word是否已開啟
'   If g_WordAp Is Nothing Then
'RestarWord:
'      Set g_WordAp = New Word.Application
'      g_WordAp.Visible = False
'   End If
'   strExc(0) = Text1 & Val(Trim(Text2)) & IIf(Text3 & Text4 = "000", "", Text3 & Text4)
'   'Modified by Lydia 2017/06/16
'   'm_TempFileName = strExc(0) & "工作通知單.doc"
'   m_TempFileName = strExc(0) & TSMailName & ".doc"
'
'   If Dir(strPath & "\" & m_TempFileName) <> "" Then
'      Kill strPath & "\" & m_TempFileName
'   End If
'   g_WordAp.Documents.Open App.path & "\" & m2FileName '範本
'   g_WordAp.ActiveDocument.SaveAs strPath & "\" & m_TempFileName
'   g_WordAp.ActiveDocument.Close
'   g_WordAp.Documents.Open strPath & "\" & m_TempFileName
'   With g_WordAp
'      .Selection.WholeStory
'      .Selection.Copy
'      'Modified by Lydia 2015/12/11 改翻譯後語系
'      'For i = 0 To 12
'      'Modified by Lydia 2016/07/07
'      'For i = 0 To 13
'      For i = 0 To 17
'         strName = ""
'         strText = ""
'         If i = 0 Then
'            strName = "外翻公司"
'            'Modified by Lydia 2018/01/04 F5588-> 外翻_舜禹
'            If FT14 = 外翻_舜禹 Then
'               strText = "江蘇舜禹翻譯"
'            'Modified by Lydia 2018/01/04 F5653-> 外翻_捷恩凱
'            ElseIf FT14 = 外翻_捷恩凱 Then
'               strText = "南京捷恩凱信息技術"
'            'Added by Lydia 2017/09/28
'            'Modified by Lydia 2018/01/04 F5698-> 外翻_迅達
'            ElseIf FT14 = 外翻_迅達 Then
'               strText = "迅達翻譯"
'            End If
'         ElseIf i = 1 Then
'            strName = "外翻聯絡1"
'             'Modified by Lydia 2018/01/04 F5588-> 外翻_舜禹
'            If FT14 = 外翻_舜禹 Then
'               'Modified by Lydia 2016/04/29 舜禹改聯絡人(5/1起)
'               'strText = "陳燕"
'               strText = "鄒甜"
'            'Modified by Lydia 2018/01/04 F5653-> 外翻_捷恩凱
'            ElseIf FT14 = 外翻_捷恩凱 Then
'               strText = "吳傑"
'            'Added by Lydia 2017/09/28
'            'Modified by Lydia 2018/01/04 F5698-> 外翻_迅達
'            ElseIf FT14 = 外翻_迅達 Then
'               strText = "薛建軍"
'            End If
'         ElseIf i = 2 Then
'            strName = "外翻電話"
'            'Modified by Lydia 2018/01/04 F5588-> 外翻_舜禹
'            If FT14 = 外翻_舜禹 Then
'               strText = "86-25-84699988" & vbCrLf & "84699966，" & vbCrLf & "84699955 # 808"
'            'Modified by Lydia 2018/01/04 F5653-> 外翻_捷恩凱
'            ElseIf FT14 = 外翻_捷恩凱 Then
'               strText = "+86 (25) 84447649"
'            'Added by Lydia 2017/09/28
'            'Modified by Lydia 2018/01/04 F5698-> 外翻_迅達
'            ElseIf FT14 = 外翻_迅達 Then
'               strText = "+86 18602509292"
'            End If
'         ElseIf i = 3 Then
'            strName = "外翻傳真"
'            'Modified by Lydia 2018/01/04 F5588-> 外翻_舜禹
'            If FT14 = 外翻_舜禹 Then
'               strText = "86-25-84699965"
'            'Modified by Lydia 2018/01/04 F5653-> 外翻_捷恩凱
'            ElseIf FT14 = 外翻_捷恩凱 Then
'               strText = "+86 (25) 84447583"
'            'Added by Lydia 2017/09/28
'            'Modified by Lydia 2018/01/04 F5698-> 外翻_迅達
'            ElseIf FT14 = 外翻_迅達 Then
'               strText = ""
'            End If
'         ElseIf i = 4 Then
'            strName = "外翻email"
'            strText = FT18
'         ElseIf i = 5 Then
'            strName = "本所案號"
'            strText = strExc(0)
'         ElseIf i = 6 Then
'            strName = "專利名稱"
'            strText = IIf(pa(5) = "", IIf(pa(6) = "", pa(7), pa(6)), pa(5))
'         ElseIf i = 7 Then
'            strName = "語系"
'            'Remove by Lydia 2018/06/25 與Sharon確認不用預設德文,有需求再自行變更
'            'strExc(1) = GetPrjNationNumber(ChangeCustomerL(pa(75)))
'            ''德文組-以代理人的國別區別
'            'If strExc(1) = "231" Then
'            '    strText = "德文"
'            'Else
'                Select Case pa(150)
'                    Case "1", "2", "4": strText = "英文" '電子電機組、化學組、機械設計組
'                    Case "3": strText = "日文"  '日文組
'                    Case Else
'                          strSql = "select na05 from fagent,nation where fa10=na01(+) and fa01='" & Mid(pa(75), 1, 8) & "' and fa02='" & Mid(pa(75), 9, 1) & "' "
'                          intI = 1
'                          Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                          If intI = 1 Then
'                             strText = "" & RsTemp(0)
'                          End If
'                End Select
'            'End If
'         ElseIf i = 8 Then
'            strName = "頁數"
'            strText = Trim(mPcnt1) & "頁"
'         ElseIf i = 9 Then
'            strName = "頁圖"
'            strText = Trim(mPcnt2)
'         ElseIf i = 10 Then
'            strName = "承辦期限"
'            If strCP48 = "" Then
'               strText = ""
'            Else
'               strText = ChangeWStringToWDateString(strCP48)
'            End If
'         ElseIf i = 11 Then
'            strName = "外翻聯絡2"
'            'Modified by Lydia 2018/01/04 F5588-> 外翻_舜禹
'            If FT14 = 外翻_舜禹 Then
'               'Modified by Lydia 2016/04/29
'               'strText = "陳燕"
'               strText = "鄒甜"
'            'Modified by Lydia 2018/01/04 F5653-> 外翻_捷恩凱
'            ElseIf FT14 = 外翻_捷恩凱 Then
'               strText = "吳傑"
'            'Added by Lydia 2017/09/28
'            'Modified by Lydia 2018/01/04 F5698-> 外翻_迅達
'            ElseIf FT14 = 外翻_迅達 Then
'               strText = "薛建軍"
'            End If
'         ElseIf i = 12 Then
'            strName = "發信日"
'            strText = ChangeWStringToWDateString(strSrvDate(1))
'         'Added by Lydia 2015/12/11
'         ElseIf i = 13 Then
'            strName = "語系2"
'            If Trim(Text1) = "FCP" Then
'               strText = "繁體中文"
'            Else
'               strText = "簡體中文"
'            End If
'         'end 2015/12/11
'         'Added by Lydia 2016/07/07
'         ElseIf i = 14 Then
'            strName = "台一聯絡人"
'            strText = "" & rsAD.Fields("st02")
'         ElseIf i = 15 Then
'            strName = "台一分機"
'            strText = PUB_chgWord2Num("" & rsAD.Fields("st07"))
'         ElseIf i = 16 Then
'            strName = "台一收信"
'            'Modified by Lydia 2017/09/28 預設為對外信箱
'            'strText = rsAD.Fields("st01") & "@taie.com.tw"
'            strText = "idept@taie.com.tw"
'         ElseIf i = 17 Then
'            strName = "台一聯絡人稱"
'            strConUser = Trim("" & rsAD.Fields("st02"))
'            If Len(strConUser) < 4 Then
'               strConUser = Left(strConUser, 1) & IIf("" & rsAD.Fields("st22") = "F", "小姐", "先生")
'            Else
'               strConUser = Left(strConUser, 2) & IIf("" & rsAD.Fields("st22") = "F", "小姐", "先生")
'            End If
'            strText = strConUser
'         'end 2016/07/07
'         End If
'         If Trim(strName) <> "" Then
'            .Selection.Find.ClearFormatting
'            .Selection.Find.Text = "|#" & strName & "#|"
'            .Selection.Find.Replacement.Text = ""
'            .Selection.Find.Forward = True
'            .Selection.Find.Wrap = wdFindContinue
'            .Selection.Find.Format = False
'            .Selection.Find.MatchCase = False
'            .Selection.Find.MatchWholeWord = False
'            .Selection.Find.MatchWildcards = False
'            .Selection.Find.MatchSoundsLike = False
'            .Selection.Find.MatchAllWordForms = False
'            .Selection.Find.MatchByte = True
'            .Selection.Find.Execute
'            .Selection.Delete
'            .Selection.TypeText strText
'         End If
'      Next i
'   End With
'
'
'   g_WordAp.ActiveDocument.Save
'   g_WordAp.ActiveDocument.Close
'
'   Clipboard.Clear '清除剪貼簿動作
'
'    '呼叫新郵件：
'    Set objOutLook = CreateObject("Outlook.Application")
'    'Modified by Lydia 2019/08/06 對外信件要加信尾(郵件範本)
'    'Set objMail = objOutLook.CreateItem(0)
'    Set objMail = objOutLook.CreateItemFromTemplate(App.path & "\$$TOT-000F22-0-01.oft")
'
'    'Modified by Lydia 2016/02/25 案號前方+張靜芳英文縮寫
'    'Modified by Lydia 2016/05/03 取消縮寫
'    'objMail.Subject = "AC/ac " & strExc(0)
'    objMail.Subject = strExc(0)
'    objMail.To = FT18
'    'Modified by Lydia 2018/01/04 F5588-> 外翻_舜禹
'    If FT14 = 外翻_舜禹 Then
'       'Modified by Lydia 2016/04/29
'       'strContent = "陳小姐"
'       strContent = "鄒小姐"
'    'Modified by Lydia 2018/01/04 F5653-> 外翻_捷恩凱
'    ElseIf FT14 = 外翻_捷恩凱 Then
'       strContent = "吳先生"
'    'Added by Lydia 2017/09/28
'    'Modified by Lydia 2018/01/04 F5698-> 外翻_迅達
'    ElseIf FT14 = 外翻_迅達 Then
'       strContent = "薛先生"
'    End If
'
'    inX = 0 'Added by Lydia 2015/12/16
'
'    'Remove by Lydia 2018/03/30 改從卷宗區抓資料
'    'Modified by Lydia 2018/04/30 P案請回復之前做法,由給舜禹翻譯資料夾帶檔案
'    If Text1.Text = "P" Then
'        'Added by Lydia 2015/11/11 將翻譯附件送入卷宗區
'        If Dir(strLoadPath, vbDirectory) <> "" Then
'            strExc(5) = Trim(Val(Text2))
'            strExc(6) = Text1 & IIf(Len(strExc(5)) < 6, "*", "") & strExc(5) & IIf(Text3 & Text4 = "000", "", Text3 & Text4) & "*.*"
'            strExc(7) = Dir(strLoadPath & "\" & strExc(6))
'            Do While strExc(7) <> ""
'                strFile = strExc(7)
'                'Added by Lydia 2015/12/22 密碼檔檔名格式：本所案號.密碼.pwd.doc ( 其中.pwd.doc不限大小寫)  例 : FCP53637.stnf011.PWD.doc
'                If strPWD = "" And InStr(UCase(strExc(7)), "PWD") > 0 Then
'                   strPWD = Mid(strExc(7), 1, InStr(UCase(strExc(7)), "PWD") - 1)
'                   strPWD = Mid(strPWD, InStr(strPWD, ".") + 1)
'                   strPWD = Replace(strPWD, ".", "")
'                End If
'                'end 2015/12/22
'
'                'Added by Lydia 2015/12/18 先檢查檔案是否開啟
'                If PUB_ChkFileOpening(strLoadPath & "\" & strFile) = True Then
'                   MsgBox strLoadPath & "\" & strFile & vbCrLf & "檔案正在使用中，請關閉或關閉檔案後間隔1分鐘，方能新增附件。", vbExclamation
'                   Exit Do
'                End If
'    'Remove by Lydia 2018/01/11 刪除-分案翻譯時自動回存原文說明書至卷宗區
'    '已和淑華確認過發Email給舜禹捷恩凱翻譯的原文說明書不一定是final版，final版可能是後面承辦上傳；同時淑華會保留寄件備份3個月，所以上傳到卷宗區已經失去意義，所以取消上傳到卷宗區的程式。
'    '            '檢查檔名規則
'    '            'Modified by Lydia 2015/12/14 副檔名改為"原文說明書"
'    '            'If PUB_ChkEmpFlowFNMRule(Text1 & "-" & Text2 & "-" & Text3 & "-" & Text4, strFile, "Y", "201", , , False, , , EMP_客戶資料) = False Then
'    '            'Modified by Lydia 2015/12/16 由靜芳自行變更副檔名
'    '            strExc(8) = ""
'    '            'If PUB_ChkEmpFlowFNMRule(Text1 & "-" & Text2 & "-" & Text3 & "-" & Text4, strFile, "Y", "201", , , False, , , "ORI") = False Then
'    '            If PUB_ChkEmpFlowFNMRule(Text1 & "-" & Text2 & "-" & Text3 & "-" & Text4, strFile, "Y", "201", , , False, , , strExc(8)) = False Then
'    '               Exit Sub
'    '            End If
'    '            'Modified by Lydia 2015/12/14 副檔名改為"原文說明書"
'    '            'If PUB_GetEmpFlowReNameFile(Text1, Text2, Text3, Text4, "201", strFile, stReName, True, 1, , , iCP09, EMP_客戶資料) = False Then Exit Sub
'    '            'Modified by Lydia 2015/12/16
'    '            'If PUB_GetEmpFlowReNameFile(Text1, Text2, Text3, Text4, "201", strFile, stReName, True, 1, , , iCP09, "ORI") = False Then Exit Sub
'    '            If PUB_GetEmpFlowReNameFile(Text1, Text2, Text3, Text4, "201", strFile, stReName, True, 1, , , iCP09, strExc(8)) = False Then Exit Sub
'    '               Set fs = CreateObject("Scripting.FileSystemObject")
'    '               Set f = fs.GetFile(strLoadPath & "\" & strExc(7))
'    '               strFile = strLoadPath & "\" & strExc(7)
'    '               '存檔
'    '               'Added by Lydia 2015/12/18 若有重複檔名,直接刪檔重傳
'    '               If IsExistCPP(iCP09, stReName) = True Then
'    '                  If DelAttFile_PDF(Text1 & "-" & Text2 & "-" & Text3 & "-" & Text4, iCP09, stReName) = False Then
'    '                     GoTo JumpSaveFile
'    '                  End If
'    '               End If
'    '               If SaveAttFile_PDF(iCP09, strFile, stReName, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), False, "A") = False Then
'    '                 'Remove by Lydia 2015/12/18
'    '                '  Exit Do
'    '               End If
'    'JumpSaveFile:
'                   strFile = strLoadPath & "\" & strExc(7)
'    'end 2018/01/11
'                   'Modified by Lydia 2016/03/10 PWD不計入附件數
'                   'inX = inX + 1 'Added by Lydia 2015/12/16 計算附件數
'                   'Added by Lydia 2015/12/22 外翻只給密碼
'                   If InStr(UCase(strExc(7)), "PWD") = 0 Then
'                       inX = inX + 1 'Added by Lydia 2016/03/10
'                       objMail.Attachments.add (strFile)
'                   End If
'                   'end 2015/12/22
'
'                   'Modified by Lydia 2015/12/18 靜芳執行時出現刪檔錯誤,有可能是因為檔案性質為唯讀
'                   'Call PUB_DelPCOrgFile(strFile) '一併將PC上的實體檔案刪除
'                   If Dir(strFile) <> "" Then
'                      SetAttr strFile, vbNormal
'                      Kill strFile
'                   End If
'                   'end 2015/12/18
'
'                strExc(7) = Dir(strLoadPath & "\" & strExc(6))
'            Loop
'        End If
'        'end 2015/11/11
'    Else '非P案
'        'Added by Lydia 2018/03/30  改從卷宗區抓資料
'        strExc(10) = ""
'        'Modified by Lydia 2018/09/18 ORI.FIX=>改成FIX.ORI
'        'Modified by Lydia 2018/10/02 +.TBL.
'        'Modified by Lydia 2018/11/30 判斷最後一道.ORI.%.PDF ,因為.FIX有人加後面
'        'strExc(1) = "SELECT CPP01,CPP02,CPP14 FROM CaseProgress A,CASEPAPERPDF B " & _
'                          "WHERE CP01='" & Text1 & "' AND CP02='" & Text2 & "' AND CP03='" & Text3 & "' AND CP04='" & Text4 & "' AND CP159=0 AND CP09=CPP01(+) " & _
'                          "AND NVL(CPP10,'N') <> 'D' AND (UPPER(CPP02) LIKE '%.ORI.PDF' OR UPPER(CPP02) LIKE '%.ORI.REP%.PDF' " & _
'                          "OR UPPER(CPP02) LIKE '%.FIX%.ORI.PDF' OR UPPER(CPP02) LIKE '%.SEQ.%' OR UPPER(CPP02) LIKE '%.PWD.%' OR UPPER(CPP02) LIKE '%.TBL.%' ) " & _
'                          "ORDER BY CPP06 DESC, CPP07 DESC "
'        'Modified by Lydia 2019/10/24 拿掉 OR UPPER(CPP02) LIKE '%.PWD.%'
'        strExc(1) = "SELECT CPP01,CPP02,CPP14 FROM CaseProgress A,CASEPAPERPDF B " & _
'                          "WHERE CP01='" & Text1 & "' AND CP02='" & Text2 & "' AND CP03='" & Text3 & "' AND CP04='" & Text4 & "' AND CP159=0 AND CP09=CPP01(+) " & _
'                          "AND NVL(CPP10,'N') <> 'D' AND ((UPPER(CPP02) LIKE '%.ORI.%' AND UPPER(CPP02) LIKE '%.PDF' ) " & _
'                          "OR UPPER(CPP02) LIKE '%.SEQ.%' OR UPPER(CPP02) LIKE '%.TBL.%' ) " & _
'                          "ORDER BY CPP06 DESC, CPP07 DESC "
'        intI = 1
'        Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
'        If intI = 1 Then
'             RsTemp.MoveFirst
'             Do While Not RsTemp.EOF
'                  If "" & RsTemp.Fields("CPP01") <> "" And "" & RsTemp.Fields("CPP02") <> "" And "" & RsTemp.Fields("CPP14") <> "" Then
'                      '說明書
'                      strFile = strPath & "\" & RsTemp.Fields("CPP02")
'                      'Modified by Lydia 2018/09/18 ORI.FIX=>改成FIX ; ORI.REP => 改成REP
'                      'Modified by Lydia 2018/11/30 判斷最後一道.ORI.%.PDF ,因為.FIX有人加後面
'                      'If InStr(UCase("" & RsTemp.Fields("CPP02")), ".ORI.") > 0 And InStr(UCase(strExc(10)), ".ORI.PDF") = 0 And _
'                                   InStr(UCase(strExc(10)), ".REP") = 0 And InStr(UCase(strExc(10)), ".FIX") = 0 Then
'                      If InStr(UCase("" & RsTemp.Fields("CPP02")), ".ORI.") > 0 And InStr(UCase(strExc(10)), ".ORI.") = 0 Then
'                             If PUB_GetFtpFile("" & RsTemp.Fields("CPP14"), strFile) = True Then
'                                 objMail.Attachments.add (strFile)
'                                 inX = inX + 1
'                                 strExc(10) = strExc(10) & strFile & "&"
'                             End If
'                      End If
'                      '序列表
'                      If InStr(UCase("" & RsTemp.Fields("CPP02")), ".SEQ.") > 0 And InStr(UCase(strExc(10)), ".SEQ.") = 0 Then
'                             If PUB_GetFtpFile("" & RsTemp.Fields("CPP14"), strFile) = True Then
'                                 objMail.Attachments.add (strFile)
'                                 inX = inX + 1
'                                 strExc(10) = strExc(10) & strFile & "&"
'                             End If
'                      End If
'                      '密碼檔
'                      If InStr(UCase("" & RsTemp.Fields("CPP02")), ".PWD.") > 0 And InStr(UCase(strExc(10)), ".PWD.") = 0 Then
'                             If PUB_GetFtpFile("" & RsTemp.Fields("CPP14"), strFile) = True Then
'                                objMail.Attachments.add (strFile)
'                                strPWD = "" & RsTemp.Fields("CPP02")
'                             End If
'                      End If
'                      'Added by Lydia 2018/10/02 需提供外翻非說明書部分之其他檔案,例如:技術用語對照表
'                      If InStr(UCase("" & RsTemp.Fields("CPP02")), "TBL.") > 0 And InStr(UCase(strExc(10)), "TBL.") = 0 Then
'                             If PUB_GetFtpFile("" & RsTemp.Fields("CPP14"), strFile) = True Then
'                                 objMail.Attachments.add (strFile)
'                                 inX = inX + 1
'                                 strExc(10) = strExc(10) & strFile & "&"
'                             End If
'                      End If
'                      'end 2018/10/02
'                  End If
'                  RsTemp.MoveNext
'             Loop
'        End If
'        If InStr(UCase(strExc(10)), ".ORI.") = 0 Or strExc(10) = "" Then
'              MsgBox "卷宗區無說明書！", vbCritical
'        End If
'        'end 2018/03/30
'    End If 'end 2018/04/30
'
'    'Move by Lydia 2015/12/16 為了計算附件數,從上方移下來
'    strContent = strContent & ", 您好:" & vbCrLf
'
'    'Added by Lydia 2015/12/22 外翻只給密碼
'    'Remove by Lydia 2019/09/27 刪除原本email內文帶入密碼之設定。
'    'If strPWD <> "" Then
'    '   'Modified by Lydia 2018/03/30 改成提醒
'    '   'strContent = "pw: " & strPWD & vbCrLf & vbCrLf & vbCrLf & strContent
'    '   strContent = "pw: (請參考附件: " & strPWD & ") " & vbCrLf & vbCrLf & vbCrLf & strContent
'    'End If
'    'end 2019/09/27
'
'    'Modified by Lydia 2015/12/16
'    'strContent = strContent & "附上" & strExc(0) & "案件電子檔共1個及工作通知單一份," & vbCrLf
'    'Modified by Lydia 2017/06/16
'    'strContent = strContent & "附上" & strExc(0) & "案件電子檔共" & inX & "個及工作通知單一份," & vbCrLf
'    strContent = strContent & "附上" & strExc(0) & "案件電子檔共" & inX & "個及" & TSMailName & "一份," & vbCrLf
'    strContent = strContent & "請確認翻譯字數及完成日期,謝謝。" & vbCrLf & vbCrLf
'    'Modified by Lydia 2016/07/07
'    'strContent = strContent & "台一張小姐"
'    strContent = strContent & "台一" & strConUser
'    '轉HTML格式
'    strContent = Replace(strContent, "新細明體", "Times New Roman")
'    strContent = Replace(strContent, vbCrLf, "<BR>")
'    strContent = Replace(strContent, "  ", "&nbsp;&nbsp;")
'    objMail.HTMLBody = "<FONT FACE=""Times New Roman"">" & strContent & "<BR>" & Replace(objMail.HTMLBody, "&lt;LetterMemo&gt;", IIf(ExceptFieldData("公用備註/英") <> "", "<BR>Message:<BR>" & Replace(ChgHTMLFormat(ExceptFieldData("公用備註/英")), vbCrLf, "<BR>") & "<BR><BR>", "&nbsp;")) & "</FONT>"
'    If m_TempFileName <> "" Then
'        objMail.Attachments.add (strPath & "\" & m_TempFileName)
'    End If
'
'    mPcnt1 = 0: mPcnt2 = 0
'    objMail.Display
'
'    Set rsAD = Nothing 'Added by Lydia 2016/07/07
'
'ErrHand:
'    If Err.Number = 462 Then '遠端伺服器不存在或無法使用
'       GoTo RestarWord
'    ElseIf Err.Number <> 0 Then
'         If strContent <> "" Then
'            MsgBox "開啟撰寫郵件視窗失敗，請人工作業！"
'         Else
'            MsgBox (Err.Description)
'         End If
'    End If
'
'   cmdFCMail(2).Enabled = True 'Added by Lydia 2015/12/18
'   Set objMail = Nothing
'   Set objOutLook = Nothing
'
'End Sub
''end 2015/10/30
'end 2025/03/13

''Added by Lydia 2015/12/18  判斷是否有卷宗區檔案
Private Function IsExistCPP(ByRef nCP09 As String, nCPP02 As String) As Boolean
Dim inR As Integer
Dim strS1 As String
Dim rsAS As New ADODB.Recordset
   IsExistCPP = False
   
   inR = 1
   strS1 = "select cpp02 from casepaperpdf where cpp01='" & nCP09 & "' and upper(cpp02)='" & UCase(nCPP02) & "'"
   Set rsAS = ClsLawReadRstMsg(inR, strS1) 'Added by Lydia 2016/04/12
   If inR = 1 Then
      IsExistCPP = True
   End If
End Function

'Add by Amy 2016/04/12 判斷案件各申請人地址是否與客戶檔地址不同
Private Function ChkCaseAndCusAddrNotAlike(ByRef stMsg As String) As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String
    Dim intR As Integer, j As Integer, jj As Integer
    Dim stTemp As String
    
    ChkCaseAndCusAddrNotAlike = False
    
    strQ = "Select 1 as No,cu23 as CAddr,cu24||cu25||cu26||cu27||cu28||cu102 as EAddr,cu29 as JAddr From Customer Where cu01='" & Mid(GetNewFagent(pa(23)), 1, 8) & "' and cu02='" & Mid(GetNewFagent(pa(23)), 9, 1) & "' "
    If pa(78) <> MsgText(601) Then
        strQ = strQ & "Union Select 2 as No,cu23 as CAddr,cu24||cu25||cu26||cu27||cu28||cu102 as EAddr,cu29 as JAddr From Customer Where cu01='" & Mid(GetNewFagent(pa(78)), 1, 8) & "' and cu02='" & Mid(GetNewFagent(pa(78)), 9, 1) & "' "
    End If
    If pa(79) <> MsgText(601) Then
        strQ = strQ & "Union Select 3 as No,cu23 as CAddr,cu24||cu25||cu26||cu27||cu28||cu102 as EAddr,cu29 as JAddr From Customer Where cu01='" & Mid(GetNewFagent(pa(79)), 1, 8) & "' and cu02='" & Mid(GetNewFagent(pa(79)), 9, 1) & "' "
    End If
    If pa(80) <> MsgText(601) Then
        strQ = strQ & "Union Select 4 as No,cu23 as CAddr,cu24||cu25||cu26||cu27||cu28||cu102 as EAddr,cu29 as JAddr From Customer Where cu01='" & Mid(GetNewFagent(pa(80)), 1, 8) & "' and cu02='" & Mid(GetNewFagent(pa(80)), 9, 1) & "' "
    End If
    If pa(81) <> MsgText(601) Then
        strQ = strQ & "Union Select 5 as No,cu23 as CAddr,cu24||cu25||cu26||cu27||cu28||cu102 as EAddr,cu29 as JAddr From Customer Where cu01='" & Mid(GetNewFagent(pa(81)), 1, 8) & "' and cu02='" & Mid(GetNewFagent(pa(81)), 9, 1) & "' "
    End If
    intR = 1
    Set RsQ = ClsLawReadRstMsg(intR, strQ)
    If intR = 1 Then
        For j = 0 To RsQ.RecordCount - 1
            stTemp = ""
            If j = 0 Then
                jj = 24
            Else
                jj = 81 + j
            End If
            'Modify by Amy 2016/05/13 CFT-015965 x64241020英文地址多個空白,造成比對錯誤
            If Replace(pa(jj), " ", "") <> Replace("" & RsQ.Fields("CAddr"), " ", "") Or Replace(pa(jj + 1), " ", "") <> Replace("" & RsQ.Fields("EAddr"), " ", "") Or Replace(pa(jj + 2), " ", "") <> Replace("" & RsQ.Fields("JAddr"), " ", "") Then
                If Replace(pa(jj), " ", "") <> Replace("" & RsQ.Fields("CAddr"), " ", "") Then stTemp = stTemp & "、中"
                If Replace(pa(jj + 4), " ", "") <> Replace("" & RsQ.Fields("EAddr"), " ", "") Then stTemp = stTemp & "、英"
                If Replace(pa(jj + 8), " ", "") <> Replace("" & RsQ.Fields("JAddr"), " ", "") Then stTemp = stTemp & "、日"
                stTemp = "申請人" & j + 1 & Mid(stTemp, 2) & "申請地址與與客戶目前地址不同，請注意！" & vbCrLf
                ChkCaseAndCusAddrNotAlike = True
            End If
            If stTemp <> MsgText(601) Then stMsg = stMsg & stTemp
            
            RsQ.MoveNext
        Next j
    End If
End Function

'Added by Lydia 2017/06/16 刪除工作通知單(FC翻譯案件郵件)
Private Sub ChkTransFile()
Dim mF  As String
Dim tmpArr As Variant 'Added by Lydia 2018/03/30

    mF = Dir(App.path & "\*" & TSMailName & ".*", vbNormal)
    Do While mF <> ""
       If PUB_ChkFileOpening(App.path & "\" & mF) = True Then
          MsgBox App.path & "\" & mF & vbCrLf & "檔案正在使用中，無法刪除。", vbExclamation
          Exit Do
       End If
       Kill App.path & "\" & mF    '刪除檔案
       'Modified by Lydia 2018/03/30 刪除從卷宗區下載的檔案
       'mF = Dir()
       strExc(1) = Mid(mF, 1, InStr(mF, TSMailName) - 1)
       If Dir(App.path & "\" & Mid(strExc(1), 1, 3) & "*" & Mid(strExc(1), 4) & "*.*") <> "" Then
            Kill App.path & "\" & Mid(strExc(1), 1, 3) & "*" & Mid(strExc(1), 4) & "*.*"
       End If
       mF = Dir(App.path & "\*" & TSMailName & ".*", vbNormal)
       'end 2018/03/30
    Loop
    'Added by Lydia 2018/06/25 刪除從卷宗區下載的檔案(新路徑)
    If strSavePath1 <> "" Then 'Added by Lydia 2018/06/26 排除無路徑的情況
        mF = Dir(strSavePath1 & "\*.*", vbNormal)
        Do While mF <> ""
           If PUB_ChkFileOpening(strSavePath1 & "\" & mF) = True Then
              MsgBox App.path & "\" & mF & vbCrLf & "檔案正在使用中，無法刪除。", vbExclamation
              Exit Do
           End If
           Kill strSavePath1 & "\" & mF    '刪除檔案
           mF = Dir(strSavePath1 & "\*.*", vbNormal)
        Loop
    End If
    'end 2018/06/25
End Sub

'Added by Morgan 2018/5/30 從 WordChinese 抽出
'美國IDS檢查
Private Function fnUsIdsChk(rsQuery As ADODB.Recordset) As Boolean
   Dim StrSQLa As String, intQ As Integer, stUSNo As String
      
   'Modified by Morgan 2018/5/30 分割母案有美國關聯案也要--禧佩
   'StrSQLa = "SELECT CR05||'-'||CR06||DECODE(CR07,'0',NULL,'-'||CR07||'-'||CR08) CR101NO,DECODE(SUBSTR(PA91,INSTR(PA91,'個體')-1,1),'大','貳萬陸仟元','小','貳萬參仟元','微','貳萬壹仟元','') FEE FROM CASERELATION,PATENT" & _
             " WHERE CR01='" & pa(1) & "' AND CR02='" & pa(2) & "' AND CR03='" & pa(3) & "' AND CR04='" & pa(4) & "'" & _
             " AND CR05=PA01(+) AND CR06=PA02(+) AND CR07=PA03(+) AND CR08=PA04(+) AND PA09='101' "
   'Modified by Morgan 2018/7/24 +判斷美國要有發明申請發文--郭 Ex.CFP-30057 -> CFP-29677, CFP-29677-1
   'Modified by Morgan 2021/3/25 改用公用函數
   'StrSQLa = "SELECT CR05||'-'||CR06||DECODE(CR07,'0',NULL,'-'||CR07||'-'||CR08) CR101NO,DECODE(SUBSTR(PA91,INSTR(PA91,'個體')-1,1),'大','貳萬陸仟元','小','貳萬參仟元','微','貳萬壹仟元','') FEE FROM CASERELATION,PATENT" & _
            " WHERE (CR01,CR02,CR03,CR04) IN (SELECT '" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "' FROM DUAL" & _
            " UNION ALL SELECT DC05,DC06,DC07,DC08 FROM DIVISIONCASE WHERE DC01='" & pa(1) & "' AND DC02='" & pa(2) & "' AND DC03='" & pa(3) & "' AND DC04='" & pa(4) & "')" & _
            " AND CR05=PA01(+) AND CR06=PA02(+) AND CR07=PA03(+) AND CR08=PA04(+) AND PA09='101' and pa08='1' and pa57 is null" & _
            " and exists(select * from caseprogress where  CP01 = PA01 And cp02 = pa02" & _
            " And cp03 = pa03 And cp04 = pa04 and cp10='101' and cp27>0 and cp159=0)" & _
            " and not exists(select * from caseprogress Where CP01 = PA01 And cp02 = pa02" & _
            " And cp03 = pa03 And cp04 = pa04 and cp10='601' and cp27>0 and cp159=0)"
            
   stUSNo = PUB_GetUSCaseNo(pa(1), pa(2), pa(3), pa(4))
   If stUSNo <> "" Then
      StrSQLa = "select '" & stUSNo & "' CR101NO from dual"
   'end 2021/3/25
      
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, StrSQLa)
      If intQ = 1 Then
         fnUsIdsChk = True
      End If
      
   End If
End Function
'Modified by Morgan 2019/2/14 +pNo:收文號
'Modfiied by Morgan 2021/12/2 +pECust:全E化客戶
Private Function GetLetterNo(Optional pNo As String, Optional pECust As Boolean = False) As String
   Dim strLetterNo As String
   
   strLetterNo = "　　　　"
   'P台灣案舉發及舉發答辯審定書之內部收文分析自動要帶來函收文號為發文字號
   'Modified by Morgan 2016/7/27 P案全面電子化,都要帶出發文號
   'If pa(1) = "P" And pa(9) = "000" Then
   '   'Modified by Morgan 2015/6/25 +501訴願,505參加訴願(目前沒定稿,先用案件性質判斷)
   '   If m_Combo8 = "16" Or Right(Combo8, 4) = "941 " Then
   If pa(1) = "P" Then
         strExc(0) = "select a.cp43,c.cp10 from CaseProgress a,CaseProgress b,CaseProgress c where a.cp01='" & pa(1) & "' and a.cp02='" & pa(2) & "' and a.cp03='" & pa(3) & "' and a.cp04='" & pa(4) & "' and a.cp10='941' and a.cp27 is null and a.cp09>'B' and a.cp43>'C'" & _
            " and b.cp09(+)=a.cp43 and b.cp10 in ('1001','1002','1503') and b.cp27 is null and c.cp09(+)=b.cp43 and c.cp10 in ('803','804','501','505')"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strLetterNo = Right(RsTemp(0), 6)
            pNo = RsTemp(0) 'Added by Morgan 2019/2/14
         'Added by Morgan 2015/2/10
         ElseIf Combo8.ItemData(Combo8.ListIndex) > 0 Then
            'Modified by Morgan 2016/3/16 改抓ItemData不再另外記錄
            'strLetterNo = Right(m_DispNum, 6)
            strLetterNo = Right(PUB_Num2DocNo(Combo8.ItemData(Combo8.ListIndex)), 6)
            pNo = PUB_Num2DocNo(Combo8.ItemData(Combo8.ListIndex)) 'Added by Morgan 2019/2/14
         'end 2015/2/10
         End If
   '   End If
   'end 2016/7/27
   End If
   'end 2014/12/8
   
   'Added by Morgan 2018/9/17 CFP電子化
   If pa(1) = "CFP" And Val(strSrvDate(1)) >= CFP第一階段電子化啟用日 Then
      If Combo8.ItemData(Combo8.ListIndex) > 0 Then 'Added by Morgan 2018/10/12
         strLetterNo = Right(PUB_Num2DocNo(Combo8.ItemData(Combo8.ListIndex)), 6)
         pNo = PUB_Num2DocNo(Combo8.ItemData(Combo8.ListIndex)) 'Added by Morgan 2019/2/14
      End If
   End If
   'end 2018/9/17
   
   'Added by Morgan 2018/10/16
   '專利處人員操作時若沒有發文號時抓該人員承辦的最後一道C類未發文程序
   'Modified by Morgan 2021/12/2 +全E化客戶
   If Trim(strLetterNo) = "" And (Left(Pub_StrUserSt03, 2) = "P1" Or pECust = True) Then
      strSql = "SELECT CP09 FROM CASEPROGRESS WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP09 LIKE 'C%' AND CP158=0 AND CP159=0 and CP14='" & strUserNum & "' order by CP05 desc,CP09 desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strLetterNo = Right(RsTemp.Fields(0), 6)
         pNo = RsTemp.Fields(0) 'Added by Morgan 2019/2/14
      End If
   Else
   'end 2018/10/16
   
      'add by sonia 2018/9/26 內商承辦之C類
      strSql = "SELECT cp14 FROM CaseProgress,staff WHERE CP09='" & PUB_Num2DocNo(Combo8.ItemData(Combo8.ListIndex)) & "' AND CP09 LIKE 'C%' AND CP14=ST01(+) AND ST03 LIKE 'P2%'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strLetterNo = Right(PUB_Num2DocNo(Combo8.ItemData(Combo8.ListIndex)), 6)
         pNo = PUB_Num2DocNo(Combo8.ItemData(Combo8.ListIndex)) 'Added by Morgan 2019/2/14
      'add by sonia 2018/10/2 操作人員為商標處時,若有C類未發文則抓該筆進度T-104135
      ElseIf Left(Pub_StrUserSt03, 2) = "P2" Then
         strSql = "SELECT SUBSTR(MAX(CP05||CP09),9) FROM CASEPROGRESS,STAFF WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP09 LIKE 'C%' AND CP158=0 AND CP159=0 AND CP14=ST01(+) AND ST03 LIKE 'P2%'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If "" & RsTemp.Fields(0) <> "" Then
               strLetterNo = Right(RsTemp.Fields(0), 6)
               pNo = RsTemp.Fields(0) 'Added by Morgan 2019/2/14
            End If
         End If
      'end 2018/10/2
      End If
      
   End If 'Added by Morgan 2018/10/16
   
   GetLetterNo = strLetterNo
End Function
'Added by Morgan 2019/6/28
'案件美金請款總額
Private Function GetBillUSAmount(pCP01 As String, pCP02 As String, pCP03 As String, pCP04 As String) As String
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   stSQL = "select a1k18,sum(a1k08 - nvl(a1k31, 0)) as Namount" & _
      " from acc1k0 where (a1k12 is null or a1k12 = 0) and a1k25 is null and a1k13 = '" & pCP01 & "'" & _
      " and a1k14 = '" & pCP02 & "' and a1k15 = '" & pCP03 & "' and a1k16 = '" & pCP04 & "' group by a1k18"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      If "" & rsQuery.Fields("a1k18") <> "USD" Or rsQuery.RecordCount > 1 Then
         MsgBox "有非美金的請款單，請款總額將不列出！", vbCritical
      Else
         GetBillUSAmount = Format(rsQuery.Fields(1), "#,###")
      End If
   End If
   Set rsQuery = Nothing
   
End Function

'add by sonia 2019/11/19
Private Function DateType(ByRef nCountry As String, nCustNo As String, nFagentNo As String) As Boolean
Dim stSQL As String, intQ As Integer
Dim rsQuery As ADODB.Recordset
   
   DateType = True
   If nCountry = "000" And nFagentNo = "" Then
      stSQL = "select cu87 from customer WHERE " & ChgCustomer(nCustNo)
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 1 Then
         If "" & rsQuery.Fields("cu87") < "010" Then
            DateType = False
         End If
      End If
      Set rsQuery = Nothing
   End If
End Function
'end 2019/11/19

'Added by Lydia 2020/07/16 取得進度檔之A類最新承辦人,案源之介紹人
Private Function GetCaseCP14(ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, Optional ByRef pLOS04List As String, Optional ByRef pLOS04n As String) As String
Dim stSQL As String, intQ As Integer
Dim rsQuery As ADODB.Recordset

    GetCaseCP14 = strUserNum '預設
    pLOS04n = ""
    pLOS04List = ""
    '進度檔之最新承辦人
    stSQL = "select cp14 from caseprogress,staff where cp01='" & pCP01 & "' and cp02= '" & pCP02 & "' and cp03= '" & pCP03 & "' and cp04= '" & pCP04 & "' " & _
                "and cp159=0 and nvl(cp14,'N') <> 'N' and cp14=st01(+) and st04='1' and substr(cp09,1,1)='A' order by cp05 desc, cp27 desc "
    intQ = 1
    Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
    If intQ = 1 Then
       GetCaseCP14 = "" & rsQuery.Fields("cp14")
    End If
    '案源之介紹人(最早)
    stSQL = "select los04,getstaffnamelist(los04) as los04n from caseprogress,LawOfficeSource where cp01='" & pCP01 & "' and cp02= '" & pCP02 & "' and cp03= '" & pCP03 & "' and cp04= '" & pCP04 & "' " & _
                 "and cp159=0 and cp162=los15(+) and nvl(los07,0)=0 and los04 is not null order by los15 "
    intQ = 1
    Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
    If intQ = 1 Then
       pLOS04List = "" & rsQuery.Fields("los04")
       pLOS04n = "" & rsQuery.Fields("los04n")
    End If
    Set rsQuery = Nothing
End Function

'Added by Lydia 2020/09/11 抓申請人1的編號
'Memo by Lydia 2021/02/08 改抓申請人1~5的編號、名稱、備註
Private Sub GetStrCustomer()
Dim intQ As Integer 'Added by Lydia 2021/02/08

    If pa(1) <> "" And pa(2) <> "" Then
        Select Case Text1.Text
            Case "P", "CFP", "FCP" '專利
                m_CustNo(1) = pa(26)
                'Added by Lydia 2021/02/08
                m_CustNo(2) = pa(27)
                m_CustNo(3) = pa(28)
                m_CustNo(4) = pa(29)
                m_CustNo(5) = pa(30)
                'end 2021/02/08
            Case "T", "TF", "CFT", "FCT" '商標
                m_CustNo(1) = pa(23)
                'Added by Lydia 2021/02/08
                m_CustNo(2) = pa(78)
                m_CustNo(3) = pa(79)
                m_CustNo(4) = pa(80)
                m_CustNo(5) = pa(81)
                'end 2021/02/08
            Case "L", "FCL", "CFL", "LIN", "ACS"  '法務
                m_CustNo(1) = pa(11)
                'Added by Lydia 2021/02/08
                m_CustNo(2) = pa(43)
                m_CustNo(3) = pa(44)
                m_CustNo(4) = pa(45)
                m_CustNo(5) = pa(46)
                'end 2021/02/08
            Case "LA" '顧問
                m_CustNo(1) = pa(5)
                'Added by Lydia 2021/02/08
                m_CustNo(2) = pa(24)
                m_CustNo(3) = pa(25)
                m_CustNo(4) = pa(26)
                m_CustNo(5) = pa(27)
                'end 2021/02/08
            Case Else '服務業務
                m_CustNo(1) = pa(8)
                'Added by Lydia 2021/02/08
                m_CustNo(2) = pa(58)
                m_CustNo(3) = pa(59)
                m_CustNo(4) = pa(65)
                m_CustNo(5) = pa(66)
                'end 2021/02/08
        End Select
        'Added by Lydia 2021/02/08 抓申請人1~5的編號+名稱(英->中->日)
        For intI = 1 To 5
            m_CustName(intI) = ""
            m_CustMemo(intI) = ""
            If Trim(m_CustNo(intI)) <> "" Then
                m_CustNo(intI) = ChangeCustomerL(m_CustNo(intI))
                strExc(0) = "select nvl(cu05,nvl(cu04,cu06)) as cname, cu79 as cmemo from customer where cu01='" & Mid(m_CustNo(intI), 1, 8) & "' and cu02='" & Mid(m_CustNo(intI), 9, 1) & "' "
                intQ = 1
                Set RsTemp = ClsLawReadRstMsg(intQ, strExc(0))
                If intQ = 1 Then
                    m_CustName(intI) = m_CustNo(intI) & " " & RsTemp.Fields("cname")
                    m_CustMemo(intI) = "" & RsTemp.Fields("cmemo")
                End If
            End If
        Next intI
        'end 2021/02/08
    End If
End Sub

'Added by Morgan 2020/12/30
'Modified by Morgan 2023/12/12 調整內容--郭雅娟
'IDS指示信
Private Sub InsIDSContent(pCP09 As String)
   Dim oTable As Word.Table
   
   g_WordAp.Visible = True
   With g_WordAp.Selection
   'Added by Moprgan 2023/12/12 從後面移來--郭
   .TypeText "Please file the IDS document with the USPTO before "
   .Font.ColorIndex = wdRed
   .TypeText "xx xx, " & Left(strSrvDate(1), 4)
   .Font.ColorIndex = wdAuto
   'Modified by Morgan 2023/12/19 --郭
   '.TypeText ", without transferring the case. When you have completed filing the IDS document with the USPTO, we would appreciate your sending us a copy of the same. If you are unable to file IDS for this case, please let us know what went wrong and kindly prevent from incurring expenses. The official filing receipt for this case is attached for your reference."
   .TypeText ", without transferring the case. When you have completed filing the IDS document with the USPTO, we would appreciate your sending us a copy of the same. If you are unable to file IDS for this case, please let us know what went wrong and kindly avoid incurring expenses. The official filing receipt for this case is attached for your reference."
   'end 2023/12/19
   .TypeParagraph
   .TypeParagraph
'Removed by Morgan 2023/12/19 取消，改回後面並用原先的內容--郭
'   .Font.ColorIndex = wdRed
'   .TypeText "【註：如果本案已收到第一次OA且他國OA未超過3個月，請加入以下段落】"
'   .Font.ColorIndex = wdAuto
'   .TypeParagraph
'   .TypeText "Please file the IDS document with the statement specified in paragraph (e). Since the issuance of "
'   .Font.ColorIndex = wdRed
'   .TypeText "search report/Office Action"
'   .Font.ColorIndex = wdAuto
'   .TypeText " is no more than three months prior to the filing of the IDS, please do not pay the official fee as set forth in § 1.17(p)."
'   .TypeParagraph
'   .TypeParagraph
'   .Font.ColorIndex = wdRed
'   .TypeText "【註：如果這個IDS是第二階段的費用，例如本案已收到第一次OA且他國OA已超過3個月，請加入以下段落】"
'   .Font.ColorIndex = wdAuto
'   .TypeParagraph
'   .TypeText "Please file the IDS document and pay the fee set forth in § 1.17(p)."
'   .TypeParagraph
'   .TypeParagraph
'   .Font.ColorIndex = wdRed
'   .TypeText "【註：如果本案已核准，但尚未領證，請加入以下段落】"
'   .Font.ColorIndex = wdAuto
'   .TypeParagraph
'   .TypeText "Please file the IDS document with the statement specified in paragraph (e) and pay the fee set forth in § 1.17(p)."
'   .TypeParagraph
'   .TypeParagraph
'end 2023/12/19
   'end 2023/12/12
   
   .TypeText "The applicant has received "
   .Font.ColorIndex = wdRed
   .TypeText "a search report/an Office Action of the related xxxx application NO. xxxxxxx mailed on xx xx, 2020.【註：紅字部分請工程師依個案情況填寫近期收到OA的國家、申請案號、OA日期】"
   .Font.ColorIndex = wdAuto
   .TypeText "The applicant authorizes you to file an IDS document as shown in the following table with the USPTO in accordance with CFR 1.56 to comply with the IDS obligation. Please find enclosed are copies of "
   .Font.ColorIndex = wdRed
   .TypeText "non-US patent, office action, and their English abstracts from EPO website as brief explanations,"
   .Font.ColorIndex = wdAuto
   .TypeText " which you can use to prepare the documents for the IDS."
   .Font.ColorIndex = wdRed
   .TypeText "【註：紅字部分請工程師依個案情況確認呈報資料說明】"
   .Font.ColorIndex = wdAuto
   .TypeParagraph
   .TypeParagraph
   
   strExc(0) = "select * from IDSlist where IL01='" & pCP09 & "' order by IL02,IL03"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Set oTable = .Tables.add(Range:=.Range, NumRows:=1, NumColumns:=1)
      
      With oTable
      '格線
      .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
      .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
      .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
      .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
      .Borders(wdBorderVertical).LineStyle = wdLineStyleSingle
      .Borders(wdBorderHorizontal).LineStyle = wdLineStyleSingle
      End With
      
      oTable.Select
      .ParagraphFormat.LeftIndent = g_WordAp.CentimetersToPoints(0.16)
      .Cells(1).SetHeight RowHeight:=16, HeightRule:=wdRowHeightAtLeast
      .Cells.VerticalAlignment = wdCellAlignVerticalCenter
      .InsertRows 2
      
      oTable.Rows(1).Cells(1).Select
      .Font.Bold = True
      .TypeText "US Patent Document"
      oTable.Rows(2).Cells(1).Select
      .Font.Bold = True
      .TypeText "Foreign Patent Document"
      oTable.Rows(3).Cells(1).Select
      .Font.Bold = True
      .TypeText "Non Patent Literature Document"
            
      oTable.Rows(1).Select
      .Collapse Direction:=wdCollapseEnd
      .InsertRows 1
      .Font.Bold = False
      .Cells.Split NumRows:=1, NumColumns:=4, MergeBeforeSplit:=True
      .Borders(wdBorderVertical).LineStyle = wdLineStyleSingle
      .Cells(1).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(1), RulerStyle:=wdAdjustProportional
      .Cells(2).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(4.55), RulerStyle:=wdAdjustProportional
      .Cells(3).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(4.25), RulerStyle:=wdAdjustProportional
      
      .Collapse Direction:=wdCollapseStart
      .TypeText "No."
      .MoveRight Unit:=wdCharacter, Count:=1
      .TypeText "Document No."
      .MoveRight Unit:=wdCharacter, Count:=1
      .TypeText "Issue/Publication date"
      .MoveRight Unit:=wdCharacter, Count:=1
      .MoveRight Unit:=wdCharacter, Count:=1
      .InsertRows 1
      .Collapse Direction:=wdCollapseStart
      RsTemp.MoveFirst
      RsTemp.Find "IL02='1'"
      If RsTemp.EOF Then
         .TypeText "1"
         .MoveRight Unit:=wdCharacter, Count:=1
         .TypeText "--"
         .MoveRight Unit:=wdCharacter, Count:=1
         .TypeText "--"
      Else
         Do While Not RsTemp.EOF
            If RsTemp("IL02") = "1" Then
               If RsTemp("IL03") > 1 Then
                  .InsertRows 1
               End If
               .TypeText RsTemp("IL03")
               .MoveRight Unit:=wdCharacter, Count:=1
               .TypeText RsTemp("IL04")
               .MoveRight Unit:=wdCharacter, Count:=1
               .TypeText Format(RsTemp("IL06"), "@@@@-@@-@@")
               .MoveRight Unit:=wdCharacter, Count:=1
               .MoveRight Unit:=wdCharacter, Count:=1
            Else
               Exit Do
            End If
            RsTemp.MoveNext
         Loop
      End If
      
      .MoveDown Unit:=wdLine, Count:=1
      .SelectRow
      .Collapse Direction:=wdCollapseEnd
      .InsertRows 1
      .Font.Bold = False
      .Cells.Split NumRows:=1, NumColumns:=5, MergeBeforeSplit:=True
      .Borders(wdBorderVertical).LineStyle = wdLineStyleSingle
      .Cells(1).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(1), RulerStyle:=wdAdjustProportional
      .Cells(2).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(4.55), RulerStyle:=wdAdjustProportional
      .Cells(3).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
      .Cells(4).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(3.5), RulerStyle:=wdAdjustProportional
      
      .Collapse Direction:=wdCollapseStart
      
      .TypeText "No."
      .MoveRight Unit:=wdCharacter, Count:=1
      .TypeText "Document No."
      .MoveRight Unit:=wdCharacter, Count:=1
      .TypeText "Country code"
      .MoveRight Unit:=wdCharacter, Count:=1
      .TypeText "Publication date"
      .MoveRight Unit:=wdCharacter, Count:=1
      .TypeText "English brief explanation"
      .MoveRight Unit:=wdCharacter, Count:=1
      .InsertRows 1
      .Collapse Direction:=wdCollapseStart
      
      RsTemp.MoveFirst
      RsTemp.Find "IL02='2'"
      If RsTemp.EOF Then
         .TypeText "1"
         .MoveRight Unit:=wdCharacter, Count:=1
         .TypeText "--"
         .MoveRight Unit:=wdCharacter, Count:=1
         .TypeText "--"
         .MoveRight Unit:=wdCharacter, Count:=1
         .TypeText "--"
         .MoveRight Unit:=wdCharacter, Count:=1
         .TypeText "--"
      Else
         Do While Not RsTemp.EOF
            If RsTemp("IL02") = "2" Then
               If RsTemp("IL03") > 1 Then
                  .InsertRows 1
               End If
               .TypeText RsTemp("IL03")
               .MoveRight Unit:=wdCharacter, Count:=1
               .TypeText RsTemp("IL04")
               .MoveRight Unit:=wdCharacter, Count:=1
               .TypeText RsTemp("IL05")
               .MoveRight Unit:=wdCharacter, Count:=1
               .TypeText Format(RsTemp("IL06"), "@@@@-@@-@@")
               .MoveRight Unit:=wdCharacter, Count:=1
               If RsTemp("IL07") = "Y" Then
                  .TypeText "Yes"
               ElseIf RsTemp("IL07") = "N" Then
                  .TypeText "No"
               End If
               .MoveRight Unit:=wdCharacter, Count:=1
            Else
               Exit Do
            End If
            RsTemp.MoveNext
         Loop
      End If
      
      .MoveDown Unit:=wdLine, Count:=1
      .SelectRow
      .Collapse Direction:=wdCollapseEnd
      .InsertRows 1
      .Font.Bold = False
      .Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=True
      .Borders(wdBorderVertical).LineStyle = wdLineStyleSingle
      .Cells(1).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(1), RulerStyle:=wdAdjustProportional
      .Cells(2).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(9), RulerStyle:=wdAdjustProportional
      
      .Collapse Direction:=wdCollapseStart
      .TypeText "No."
      .MoveRight Unit:=wdCharacter, Count:=1
      .TypeText "Author, title, date, or country where published"
      .MoveRight Unit:=wdCharacter, Count:=1
      .TypeText "English brief explanation"
      .MoveRight Unit:=wdCharacter, Count:=1
      .InsertRows 1
      
      RsTemp.MoveFirst
      RsTemp.Find "IL02='3'"
      If RsTemp.EOF Then
         .TypeText "1"
         .MoveRight Unit:=wdCharacter, Count:=1
         .TypeText "--"
         .MoveRight Unit:=wdCharacter, Count:=1
         .TypeText "--"
      Else
         Do While Not RsTemp.EOF
            If RsTemp("IL02") = "3" Then
               If RsTemp("IL03") > 1 Then
                  .InsertRows 1
               End If
               .TypeText RsTemp("IL03")
               .MoveRight Unit:=wdCharacter, Count:=1
               .TypeText RsTemp("IL04")
               .MoveRight Unit:=wdCharacter, Count:=1
               If RsTemp("IL07") = "Y" Then
                  .TypeText "Yes"
               ElseIf RsTemp("IL07") = "N" Then
                  .TypeText "No"
               End If
               .MoveRight Unit:=wdCharacter, Count:=1
            Else
               Exit Do
            End If
            RsTemp.MoveNext
         Loop
      End If
   End If
   .EndKey Unit:=wdStory
   .TypeParagraph
   'Added by Morgan 2024/1/26--郭
   .Font.ColorIndex = wdRed
   .TypeText "【註：如果本案尚未收到第一次OA，請加入以下段落】"
   .Font.ColorIndex = wdAuto
   .TypeParagraph
   .TypeText "Please file the IDS document without submitting the statement specified in paragraph (e) and without paying the fee set forth in § 1.17(p), since the referenced patent application has not received Office Action."
   'Removed by Morgan 2024/3/26 多了
   '.Font.ColorIndex = wdRed
   '.TypeText "search report/Office Action"
   '.Font.ColorIndex = wdAuto
   'end 2024/3/26
   .TypeParagraph
   .TypeParagraph
   'end 2024/1/26
   'Added by Morgan 2023/5/10 --李柏翰
   'Removed by Morgan 2023/12/13 移到第2段並修改內容--郭
   'Added by Morgan 2023/12/19 又從上面移回來並用原先的內容--郭
   .Font.ColorIndex = wdRed
   .TypeText "【註：如果本案已收到第一次OA且他國OA未超過3個月，請加入以下段落】"
   .Font.ColorIndex = wdAuto
   .TypeParagraph
   .TypeText "Please file the IDS document with the statement specified in paragraph (e) rather than pay the fee set forth in § 1.17(p), since the "
   .Font.ColorIndex = wdRed
   .TypeText "search report/Office Action"
   .Font.ColorIndex = wdAuto
   .TypeText " is not more than three months prior to the filing of the IDS."
   .TypeParagraph
   .TypeParagraph
   .Font.ColorIndex = wdRed
   .TypeText "【註：如果這個IDS是第二階段的費用，例如本案已收到第一次OA且他國OA已超過3個月，請加入以下段落】"
   .Font.ColorIndex = wdAuto
   .TypeParagraph
   .TypeText "Please file the IDS document and pay the fee set forth in § 1.17(p)."
   .TypeParagraph
   .TypeParagraph
   .Font.ColorIndex = wdRed
   .TypeText "【註：如果本案已核准，但尚未領證，請加入以下段落】"
   .Font.ColorIndex = wdAuto
   .TypeParagraph
   .TypeText "Please file the IDS document with the statement specified in paragraph (e) and pay the fee set forth in § 1.17(p)."
   'end 2023/12/19
   'end 2023/12/13
   'end 2023/5/10
'Removed by Morgan 2023/12/13 移到第1段並修改內容--郭
'   .TypeParagraph
'   .TypeParagraph
'   .TypeText "Please file the IDS document with the USPTO before "
'   .Font.ColorIndex = wdRed
'   .TypeText "xx xx, " & Left(strSrvDate(1), 4) & "."
'   .Font.ColorIndex = wdAuto
'   .TypeText " When you have completed filing the IDS document with the USPTO, we would appreciate you sending us a copy of the same."
'end 2023/12/13
   End With
End Sub

'Added by Morgan 2023/9/12
'檢查是否有B類未發文
Private Function ChkBCP() As Boolean
   strExc(0) = "select CP09,CP10 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and substr(cp09,1,1)='B' and cp27||cp57 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      ChkBCP = True
   End If
End Function

'Added by Morgan 2023/12/13 從Read
'設定CF代理人相關欄位
Private Sub SetCFRef()
   Combo7.Clear
   'Modified by Morgan 2024/1/15 +FA16
   strExc(0) = "SELECT FA04,FA05,FA63,FA64,FA65,FA06,FA12,FA14,FA13,FA15,FA16 FROM FAGENT WHERE " & ChgFagent(m_strCP44)
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
       With RsTemp
         m_strCP44_FA16 = "" & .Fields("FA16")
         'Add by Morgan 2004/9/29
         m_strFax(1) = "" & .Fields("FA12")
         m_strFax(0) = "" & .Fields("FA14")
         '2004/9/29 end
         'Add by Morgan 2007/1/19
         m_strFax(3) = "" & .Fields("FA13")
         m_strFax(2) = "" & .Fields("FA15")
         'END 2007/1/19
         For i = 0 To 5
             If IsNull(.Fields(i)) = False And (.Fields(i)) <> "" Then
                 cfa(i) = .Fields(i)
                 Combo7.AddItem cfa(i)
                 'Combo7 = cfa(0) 'Removed by Morgan 2023/12/13
             End If
         Next
         If Combo7.ListCount > 0 Then Combo7.ListIndex = 0 'Added by Morgan 2023/12/13
       End With
   End If
   Combo7.Tag = m_strCP44
End Sub

'Added by Morgan 2023/12/13
'設定CFP美國發明案有發文IDS時CF代理人特殊規則
'm_strCP09目前為日文傳真封面用,美國案應可不必設定
Private Sub SetCFPUSAgent()
   Dim stIDSCP44 As String, stIDSCP45 As String
   
   'Added by Morgan 2024/1/4 若只有1道未發文且有代理人則直接帶 Ex:CFP-033012 IDS
   strExc(0) = "select CP10,CP44,CP45,CP43 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " and cp09<'C' and cp158=0 and cp159=0"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.RecordCount = 1 Then
         If Not IsNull(RsTemp("cp44")) Then
            m_strCP44 = "" & RsTemp.Fields("CP44").Value
            m_strCP45 = "" & RsTemp.Fields("CP45").Value
            Exit Sub
         End If
      End If
   End If
   'end 2024/1/4
   
   '先檢查是否有IDS發文
   'Modified by Morgan 2024/1/15 改判斷是否有收文IDS且有設定代理人
   'strExc(0) = "select CP44,CP45 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " and cp27>0 and cp159=0 and cp10='214' order by cp27 desc"
   strExc(0) = "select CP44,CP45 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " and cp159=0 and cp10='214' and cp44 is not null order by cp27 desc"
   'end 2024/1/15
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      stIDSCP44 = "" & RsTemp("cp44")
      stIDSCP45 = "" & RsTemp.Fields("CP45").Value
      '若IDS代理人與案件最新發文代理人不同
      If m_strCP44 <> stIDSCP44 Then
         '檢查是否有AB類未發文
         strExc(0) = "select CP10,CP44,CP45,CP43 FROM CaseProgress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " and cp09<'C' and cp158=0 and cp159=0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            '若只有1筆
            If RsTemp.RecordCount = 1 Then
               '若是936(回覆委任代理人)或957(詢問代理人)
               'Modified by Morgan 2024/1/15 +214
               If RsTemp("cp10") = "936" Or RsTemp("cp10") = "957" Or RsTemp("cp10") = "214" Then
                  '已設定代理人
                  If Not IsNull(RsTemp("cp44")) Then
                     m_strCP44 = "" & RsTemp.Fields("CP44").Value
                     m_strCP45 = "" & RsTemp.Fields("CP45").Value
                  '未設定代理人
                  Else
                     strExc(1) = "" & RsTemp("cp43")
                     '若有相關收文號
                     If strExc(1) <> "" Then
                        If Left(strExc(1), 1) = "C" Then
                           strExc(0) = "select b.cp44,b.cp45 from CaseProgress a,caseprogress b where a.cp09='" & strExc(1) & "' and b.cp09(+)=a.cp43 and b.cp10='214'"
                        Else
                           strExc(0) = "select * from CaseProgress where cp09='" & strExc(1) & "' and cp10='214'"
                        End If
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           m_strCP44 = "" & RsTemp("cp44")
                           m_strCP45 = "" & RsTemp.Fields("CP45").Value
                        End If
                     '若無相關收文號(內部收文及分案已管控,應該不會發生)
                     Else
                        If MsgBox("請問本次指示信的對象是否為IDS代理人？", vbYesNo + vbQuestion) = vbYes Then
                           m_strCP44 = stIDSCP44
                           m_strCP45 = stIDSCP45
                        End If
                     End If
                  End If
               End If
            '有多筆未發文且有936(回覆委任代理人)或957(詢問代理人)時詢問發信對象
            Else
               intI = 0
               RsTemp.MoveFirst
               RsTemp.Find "cp10='936'"
               If Not RsTemp.EOF Then
                  intI = 1
               Else
                  RsTemp.MoveFirst
                  RsTemp.Find "cp10='957'"
                  If Not RsTemp.EOF Then
                     intI = 1
                  'Added by Morgan 2024/1/15 +214
                  Else
                     RsTemp.MoveFirst
                     RsTemp.Find "cp10='214'"
                     If Not RsTemp.EOF Then
                        intI = 1
                     End If
                  'end 2024/1/15
                  End If
               End If
               If intI = 1 Then
                  If MsgBox("請問本次指示信的對象是否為IDS代理人？", vbYesNo + vbQuestion) = vbYes Then
                     m_strCP44 = stIDSCP44
                     m_strCP45 = stIDSCP45
                  End If
               End If
            End If
         End If
      End If
   End If
End Sub
