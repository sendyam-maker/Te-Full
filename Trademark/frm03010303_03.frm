VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03010303_03 
   BorderStyle     =   1  '單線固定
   Caption         =   "註冊證 / 延展證書輸入"
   ClientHeight    =   5170
   ClientLeft      =   2520
   ClientTop       =   2800
   ClientWidth     =   9160
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5170
   ScaleWidth      =   9160
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   15
      Top             =   4728
      Width           =   492
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6420
      MaxLength       =   8
      TabIndex        =   8
      Top             =   3735
      Width           =   1092
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5820
      MaxLength       =   8
      TabIndex        =   14
      Top             =   4380
      Width           =   1092
   End
   Begin VB.CommandButton cmdInputTG 
      Caption         =   "商品及服務資料輸入(&I)"
      Height          =   400
      Left            =   3600
      Style           =   1  '圖片外觀
      TabIndex        =   52
      Top             =   60
      Width           =   2295
   End
   Begin VB.TextBox textCP18 
      Height          =   285
      Left            =   5820
      TabIndex        =   10
      Top             =   4053
      Width           =   1092
   End
   Begin VB.TextBox TextMyanmar_3 
      Height          =   285
      Left            =   3060
      MaxLength       =   6
      TabIndex        =   13
      Top             =   4380
      Width           =   492
   End
   Begin VB.TextBox TextMyanmar_2 
      Height          =   285
      Left            =   2220
      MaxLength       =   4
      TabIndex        =   12
      Top             =   4380
      Width           =   492
   End
   Begin VB.TextBox TextMyanmar_1 
      Height          =   285
      Left            =   1260
      MaxLength       =   3
      TabIndex        =   11
      Top             =   4380
      Width           =   492
   End
   Begin VB.TextBox textSP13 
      Height          =   285
      Left            =   5820
      MaxLength       =   20
      TabIndex        =   1
      Top             =   2781
      Width           =   2532
   End
   Begin VB.TextBox textFeeDate 
      Height          =   264
      Left            =   5820
      MaxLength       =   8
      TabIndex        =   6
      Top             =   3417
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5820
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   2145
      Width           =   2532
   End
   Begin VB.TextBox textRegDate 
      Height          =   285
      Left            =   1260
      MaxLength       =   8
      TabIndex        =   2
      Top             =   3099
      Width           =   1092
   End
   Begin VB.TextBox textTM15 
      Height          =   285
      Left            =   5820
      MaxLength       =   20
      TabIndex        =   3
      Top             =   3099
      Width           =   2532
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6900
      TabIndex        =   17
      Top             =   60
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5940
      TabIndex        =   16
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8160
      TabIndex        =   18
      Top             =   60
      Width           =   912
   End
   Begin VB.TextBox textPrint 
      Height          =   285
      Left            =   1260
      MaxLength       =   1
      TabIndex        =   7
      Top             =   3735
      Width           =   732
   End
   Begin VB.TextBox textTM14 
      Height          =   285
      Left            =   1260
      MaxLength       =   8
      TabIndex        =   0
      Top             =   2781
      Width           =   1092
   End
   Begin VB.TextBox textTM22 
      Height          =   285
      Left            =   2700
      MaxLength       =   8
      TabIndex        =   5
      Top             =   3417
      Width           =   1092
   End
   Begin VB.TextBox textTM21 
      Height          =   285
      Left            =   1260
      MaxLength       =   8
      TabIndex        =   4
      Top             =   3417
      Width           =   1092
   End
   Begin VB.TextBox textFee 
      Height          =   285
      Left            =   1260
      TabIndex        =   9
      Top             =   4053
      Width           =   1092
   End
   Begin VB.TextBox textTM15S 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5820
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1827
      Width           =   2532
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2145
      Width           =   2540
   End
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5820
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   540
      Width           =   2532
   End
   Begin VB.TextBox textTM09 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1827
      Width           =   2532
   End
   Begin VB.TextBox textTM27 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5820
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1509
      Width           =   2532
   End
   Begin VB.TextBox textTM08 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1509
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   540
      Width           =   2532
   End
   Begin VB.Label Label1 
      Caption         =   "(Y:內部收文更改)"
      Height          =   264
      Index           =   1
      Left            =   2952
      TabIndex        =   60
      Top             =   4776
      Width           =   1464
   End
   Begin VB.Label Label1 
      Caption         =   "是否更改註冊證/延展證書 :"
      Height          =   264
      Index           =   6
      Left            =   180
      TabIndex        =   59
      Top             =   4740
      Width           =   2244
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   1260
      TabIndex        =   58
      Top             =   2463
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   1260
      TabIndex        =   57
      Top             =   858
      Width           =   7875
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13891;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1260
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   1185
      Width           =   7125
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "12568;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "柬埔寨 延展核准日 :"
      Height          =   180
      Left            =   4740
      TabIndex        =   55
      Top             =   3735
      Width           =   1575
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
      Left            =   3810
      TabIndex        =   54
      Top             =   570
      Width           =   645
   End
   Begin VB.Label Label16 
      Caption         =   "緬甸延展日報日期 :"
      Height          =   255
      Left            =   4200
      TabIndex        =   53
      Top             =   4380
      Width           =   1575
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "點    數 :"
      Height          =   180
      Left            =   4740
      TabIndex        =   51
      Top             =   4053
      Width           =   630
   End
   Begin VB.Label Label12 
      Caption         =   "頁"
      Height          =   255
      Left            =   3660
      TabIndex        =   50
      Top             =   4380
      Width           =   255
   End
   Begin VB.Label Label11 
      Caption         =   "卷"
      Height          =   255
      Left            =   2820
      TabIndex        =   49
      Top             =   4380
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "冊"
      Height          =   255
      Left            =   1860
      TabIndex        =   48
      Top             =   4380
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "緬甸公報 :"
      Height          =   255
      Left            =   180
      TabIndex        =   47
      Top             =   4380
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "登記號 :"
      Height          =   255
      Left            =   4740
      TabIndex        =   46
      Top             =   2781
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "繳年費期限 :"
      Height          =   255
      Left            =   4740
      TabIndex        =   45
      Top             =   3417
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "申請案號 :"
      Height          =   252
      Index           =   13
      Left            =   4740
      TabIndex        =   44
      Top             =   2145
      Width           =   852
   End
   Begin VB.Label Label7 
      Caption         =   "審定號數 :"
      Height          =   255
      Left            =   4740
      TabIndex        =   42
      Top             =   3114
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "註冊日 :"
      Height          =   255
      Left            =   180
      TabIndex        =   41
      Top             =   3114
      Width           =   855
   End
   Begin VB.Label Label23 
      Caption         =   "(N:不印)"
      Height          =   255
      Left            =   2100
      TabIndex        =   40
      Top             =   3735
      Width           =   855
   End
   Begin VB.Label Label22 
      Caption         =   "列印定稿 :"
      Height          =   255
      Left            =   180
      TabIndex        =   39
      Top             =   3735
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "公告日 :"
      Height          =   255
      Left            =   180
      TabIndex        =   38
      Top             =   2781
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   2460
      X2              =   2580
      Y1              =   3540
      Y2              =   3540
   End
   Begin VB.Label Label14 
      Caption         =   "專用期限 :"
      Height          =   255
      Left            =   180
      TabIndex        =   37
      Top             =   3417
      Width           =   975
   End
   Begin VB.Label Label18 
      Caption         =   "領證費 :"
      Height          =   252
      Left            =   180
      TabIndex        =   36
      Top             =   4053
      Width           =   852
   End
   Begin VB.Label Label2 
      Caption         =   "審定號數 :"
      Height          =   255
      Left            =   4740
      TabIndex        =   35
      Top             =   1830
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   252
      Index           =   11
      Left            =   180
      TabIndex        =   34
      Top             =   2463
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   250
      Index           =   10
      Left            =   180
      TabIndex        =   33
      Top             =   2150
      Width           =   1060
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家 :"
      Height          =   252
      Index           =   8
      Left            =   4740
      TabIndex        =   32
      Top             =   540
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "商品類別 :"
      Height          =   252
      Index           =   7
      Left            =   180
      TabIndex        =   31
      Top             =   1827
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "正商標號數 :"
      Height          =   252
      Index           =   4
      Left            =   4740
      TabIndex        =   30
      Top             =   1509
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "商標種類 :"
      Height          =   252
      Index           =   2
      Left            =   180
      TabIndex        =   29
      Top             =   1509
      Width           =   852
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   180
      TabIndex        =   28
      Top             =   1191
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   252
      Left            =   180
      TabIndex        =   27
      Top             =   858
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   180
      TabIndex        =   26
      Top             =   540
      Width           =   852
   End
End
Attribute VB_Name = "frm03010303_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/08/12 改成Form2.0 ;textTM23、cmbTM05、textCP13
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 來函收文日
Dim m_CP05 As String
' 收文號
Dim m_CP09 As String
' 原案件性質
Dim m_CP10 As String
' 原業務區
Dim m_CP12 As String
' 原智權人員代號
Dim m_CP13 As String
' 商標種類代碼
Dim m_TM08 As String
' 國家代碼
Dim m_TM10 As String
' 原專用期限起日
Dim m_TM21 As String
' 原專用期限止日
Dim m_TM22 As String
' 延展後專用期限止日
Dim New_TM22 As String
' 延展前專用期限止日
Dim Pre_TM22 As String
Dim m_TM23 As String
' 申請國家的延展年度
Dim m_NA14 As Integer
' 下一程序中的本所期限(存檔及定稿使用)
Dim m_NP08 As String
Dim m_NP22 As String 'Add By Sindy 2011/5/24
Dim strCP09 As String 'Modify By Sindy 2011/5/24
'
Dim m_SP48 As String
'
'Add By Cheng 2002/07/22
Dim m_TM13 As String '准駁通知日
'910918 nick 正商標號數
Dim m_TM27 As String
'add by nick 2004/09/16 檢查是否已經有商品及服務
Public ChkTG As Boolean
' 申請日
Dim m_TM11 As String
'add by nickc 2006/12/21  '記錄柬埔寨使用宣誓期間
Dim m_046Date2 As String
Dim strET01 As String, strET03 As String 'Add By Sindy 2023/5/3
'Add By Sindy 2023/4/27
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2023/4/27 END
Dim bolAddNa14 As Boolean 'Added by Lydia 2025/09/02

'Add By Sindy 2023/4/27
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

' 原資料是否有實際結果
Private Sub cmdCancel_Click()
   Unload Me
   frm03010303_02.Show
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2003/07/09
'move to unload by nick
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm03010303_02
   Unload frm03010303_01
   Unload Me
End Sub

Private Sub cmdInputTG_Click()
frm03010303_04.Hide
Set frm03010303_04.UpForm = Me
frm03010303_04.TGKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
frm03010303_04.AllClass = textTM09.Text
Me.Hide
frm03010303_04.QueryData
frm03010303_04.Show vbModal 'Modify By Sindy 2009/09/17 改為強制回應表單
End Sub

Private Sub cmdok_Click()
   If CheckDataValid = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
          'add by nickc 2005/04/22
          Pub_EndModCashMsg m_TM10
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 存檔
      'edit by nick 2004/11/03
      'OnSaveData
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      
      ' 列印定稿
      If textPrint <> "N" Then
         PrintLetter
      End If
      
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      'Added by Lydia 2025/09/02 若為註冊證輸入(相關總收文號是申請101)且案件曾發文緩審延展109，則提醒「此案曾發文緩審延展，請自行調整定稿內容 !」
      If bolAddNa14 = True Then MsgBox "此案曾發文緩審延展，請自行調整定稿內容 !", vbInformation
         
      'Add By Sindy 2023/4/27
      If Me.m_strIR01 <> "" Then
         Unload frm03010303_02
         Unload frm03010303_01
         If Not m_PrevForm Is Nothing Then
            Call m_PrevForm.GoNext
         End If
         Unload Me
      Else
      '2023/4/27 END
         Unload Me
         Unload frm03010303_02
         'Add By Cheng 2002/07/11
         frm03010303_01.textTM01.Text = Empty
         frm03010303_01.textTM02.Text = Empty
         frm03010303_01.textTM02_2.Text = Empty
         frm03010303_01.textTM03.Text = Empty
         frm03010303_01.textTM04.Text = Empty
         frm03010303_01.Show
      End If
   End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM08.BackColor = &H8000000F
   textTM09.BackColor = &H8000000F
   textTM10.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM15S.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   textTM27.BackColor = &H8000000F
      
   textCP05S.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   
   MoveFormToCenter Me
   
   'Add By Sindy 2023/4/27
   m_strIR01 = frm03010303_02.m_strIR01
   m_strIR02 = frm03010303_02.m_strIR02
   m_strIR03 = frm03010303_02.m_strIR03
   m_strIR04 = frm03010303_02.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2023/4/27 END
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP05 = Empty
      m_CP09 = Empty
   End If
   
   Select Case nType
      ' 本所案號 欄位1
      Case 0: m_TM01 = strData
                  'add by nick 2004/09/16
                  If UCase(m_TM01) = "CFT" Then cmdInputTG.Visible = True Else cmdInputTG.Visible = False
      ' 本所案號 欄位2
      Case 1: m_TM02 = strData
      ' 本所案號 欄位3
      Case 2: m_TM03 = strData
      ' 本所案號 欄位4
      Case 3: m_TM04 = strData
      ' 來函收文日
      Case 4: m_CP05 = strData
      ' 收文號
      Case 5: m_CP09 = strData
   End Select
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得各國家的未使用撤銷年度
' Input : strNation ==> 國家代碼
' Output : 傳回國家的未使用撤銷年度
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetNA19(ByVal strNation As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   GetNA19 = Empty
   
   strSql = "SELECT * FROM NATION " & _
            "WHERE NA01 = '" & strNation & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("NA19")) = False Then
         GetNA19 = rsTmp.Fields("NA19")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取商標基本檔
Private Sub QueryTradeMark()
   Dim strSql As String
   Dim strSub As String
   Dim rsTmp As New ADODB.Recordset
   Dim rsSub As ADODB.Recordset
   
   ' 取得商標基本檔的相關項目
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
         textTM10 = GetNationName(rsTmp.Fields("TM10"), 0)
         m_NA14 = GetNationExtentYear(rsTmp.Fields("TM10"))
      End If
      ' 商標名稱(中)
      If IsNull(rsTmp.Fields("TM05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM05")
      End If
      ' 商標名稱(英)
      If IsNull(rsTmp.Fields("TM06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM06")
      End If
      ' 商標名稱(日)
      If IsNull(rsTmp.Fields("TM07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM07")
      End If
      ' 顯示商標名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 商標種類
      If IsNull(rsTmp.Fields("TM08")) = False Then
         m_TM08 = rsTmp.Fields("TM08")
         If m_TM10 < "010" Then
            textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
         Else
            textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 1)
         End If
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"))
         m_TM23 = rsTmp.Fields("TM23")
      End If
      ' 商品類別
      If IsNull(rsTmp.Fields("TM09")) = False Then
         textTM09 = rsTmp.Fields("TM09")
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
      End If
      '2005/5/23 ADD BY SONIA 申請日
      If IsNull(rsTmp.Fields("TM11")) = False Then
         m_TM11 = rsTmp.Fields("TM11")
      End If
      '2005/5/23 END
      ' 公告日
      If IsNull(rsTmp.Fields("TM14")) = False Then
         textTM14 = DBDATE(rsTmp.Fields("TM14"))
      End If
      ' 審定號數
      If IsNull(rsTmp.Fields("TM15")) = False Then
         textTM15S = rsTmp.Fields("TM15")
         textTM15 = rsTmp.Fields("TM15")
      End If
      ' 專用期間(起)
      If IsNull(rsTmp.Fields("TM21")) = False Then
         'm_TM21 = rsTmp.Fields("TM21")
         textTM21 = DBDATE(rsTmp.Fields("TM21"))
      End If
      ' 專用期間(迄)
      If IsNull(rsTmp.Fields("TM22")) = False Then
         m_TM22 = rsTmp.Fields("TM22")
         '91.11.10 CANCEL BY SONIA
         'textTM22 = DBDATE(rsTmp.Fields("TM22"))
         '91.11.10 END
      End If
      'Add By Cheng 2002/07/22
      m_TM13 = "" & rsTmp.Fields("TM13").Value
      '910918 nick 正商標號數
      m_TM27 = CheckStr(rsTmp.Fields("tm27").Value)
      ' 註冊日  910919 nick
      If IsNull(rsTmp.Fields("TM20")) = False Then
         textRegDate = DBDATE(rsTmp.Fields("TM20"))
      End If
      'add by nickc 2006/05/29 加入閉卷提示
      If IsNull(rsTmp.Fields("TM29")) Then
         Me.lblClose.Caption = ""
      Else
         Me.lblClose.Caption = "已閉卷"
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 取得服務業務基本檔的欄位內容
Private Sub QueryServicePractice()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_TM01 & "' AND " & _
                  "SP02 = '" & m_TM02 & "' AND " & _
                  "SP03 = '" & m_TM03 & "' AND " & _
                  "SP04 = '" & m_TM04 & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
      ' 案件中文名稱
      If IsNull(rsTmp.Fields("SP05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP05")
      End If
      ' 案件英文名稱
      If IsNull(rsTmp.Fields("SP06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP06")
      End If
      ' 案件日文名稱
      If IsNull(rsTmp.Fields("SP07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP07")
      End If
      ' 顯示商標名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("SP08")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("SP08"))
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_TM10 = rsTmp.Fields("SP09")
         textTM10 = GetNationName(rsTmp.Fields("SP09"), 0)
         m_NA14 = GetNationExtentYear(rsTmp.Fields("SP09"))
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("SP11")) = False Then
         textTM12 = rsTmp.Fields("SP11")
      End If
      ' 專用期間(起)
      If IsNull(rsTmp.Fields("SP20")) = False Then
         'm_TM21 = rsTmp.Fields("SP20")
         textTM21 = DBDATE(rsTmp.Fields("SP20"))
      End If
      ' 專用期間(迄)
      If IsNull(rsTmp.Fields("SP21")) = False Then
         'textTM22 = DBDATE(rsTmp.Fields("SP21"))
      End If
      '
      If IsNull(rsTmp.Fields("SP48")) = False Then
         m_SP48 = DBDATE(rsTmp.Fields("SP48"))
      End If
      
      'add by nickc 2006/05/29 加入閉卷提示
      If IsNull(rsTmp.Fields("sp15")) Then
         Me.lblClose.Caption = ""
      Else
         Me.lblClose.Caption = "已閉卷"
      End If
      
      ' 若系統類別為CFC時, 帶出登記號及註冊號數
      textSP13 = Empty
      If m_TM01 = "CFC" Then
         ' 可讓使用者輸入登記號
         textSP13.BackColor = &H80000005
         textSP13.Enabled = True
         textSP13.TabStop = True
         If IsNull(rsTmp.Fields("SP13")) = False Then
            textSP13 = rsTmp.Fields("SP13")
         End If
         If IsNull(rsTmp.Fields("SP14")) = False Then
            textTM15 = rsTmp.Fields("SP14")
         End If
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 讀取案件進度檔
Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSubSQL As String
   Dim rsSubTmp As ADODB.Recordset
   
   ' 來函收文日
   textCP05S = m_CP05
   
   ' 取得案件進度檔檔案中欄位
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
      End If
      ' 智權人員
      Set rsSubTmp = New ADODB.Recordset
      strSubSQL = "SELECT * FROM CaseProgress " & _
                  "WHERE CP01 = '" & m_TM01 & "' AND " & _
                        "CP02 = '" & m_TM02 & "' AND " & _
                        "CP03 = '" & m_TM03 & "' AND " & _
                        "CP04 = '" & m_TM04 & "' AND " & _
                        "CP09 LIKE 'A%' AND " & _
                        "CP05 IN (SELECT MAX(CP05) FROM CaseProgress " & _
                                 "WHERE CP01 = '" & m_TM01 & "' AND " & _
                                 "CP02 = '" & m_TM02 & "' AND " & _
                                 "CP03 = '" & m_TM03 & "' AND " & _
                                 "CP04 = '" & m_TM04 & "' AND " & _
                                 "CP09 LIKE 'A%' "
      rsSubTmp.CursorLocation = adUseClient
      rsSubTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsSubTmp.RecordCount > 0 Then
         rsSubTmp.MoveFirst
         If IsNull(rsSubTmp.Fields("CP12")) = False Then
            m_CP12 = rsSubTmp.Fields("CP12")
         End If
         If IsNull(rsSubTmp.Fields("CP13")) = False Then
            m_CP13 = rsSubTmp.Fields("CP13")
            textCP13 = GetStaffName(rsSubTmp.Fields("CP13"))
         End If
      Else
         If IsNull(rsTmp.Fields("CP12")) = False Then
            m_CP12 = rsTmp.Fields("CP12")
         End If
         If IsNull(rsTmp.Fields("CP13")) = False Then
            m_CP13 = rsTmp.Fields("CP13")
            textCP13 = GetStaffName(rsTmp.Fields("CP13"), True)
         End If
      End If
      rsSubTmp.Close
      Set rsSubTmp = Nothing
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

Public Sub QueryData()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim nIndex As Integer
   
   ' 先預設登記號不可輸入
   textSP13.BackColor = &H8000000F
   textSP13.Enabled = False
   textSP13.TabStop = False
   'Add By Cheng 2002/07/22
   m_TM13 = ""
   'add by nick 2004/09/16
   If UCase(m_TM01) = "CFT" Then cmdInputTG.Visible = True Else cmdInputTG.Visible = False
   
   ' 讀取基本檔
   Select Case m_TM01
      ' 系統類別為CFT的為讀取商標基本檔
      Case "CFT":
         QueryTradeMark
      ' 系統類別為CFC的為讀取服務業務基本檔
      Case Else:
         QueryServicePractice
   End Select
   
   ' 讀取案件進度檔
   QueryCaseProgress
   
   ' 登記號
   If m_TM01 = "CFC" Then
      EnableTextBox textSP13, True
   Else
      EnableTextBox textSP13, False
   End If

   ' 申請國家為西班牙時, 繳年費期限為專用期限起日加上5年
   '2005/3/23 modify by sonia 加入分割
   'If m_TM10 = "211" And m_CP10 = "101" Then
   'Modify By Sindy 2012/9/20 西班牙已取消繳年費制度, 故請取消畫面 繳年費期限 的輸入及控制
'   If m_TM10 = "211" And (m_CP10 = "101" Or m_CP10 = "308") Then
'      'textFeeDate.BackColor = &H80000005
'      'textFeeDate.Enabled = True
'      'textFeeDate.TabStop = True
'      If IsEmptyText(textTM21) = False Then
'        'Modify By Cheng 2003/09/02
''         textFeeDate = DBDATE(DateSerial(Val(DBYEAR(textTM21)) + 5, Val(DBMONTH(textTM21)), Val(DBDAY(textTM21))))
'         textFeeDate = DBDATE(DateAdd("yyyy", 5, ChangeWStringToWDateString(DBDATE(textTM21))))
'      End If
'   Else
'      'textFeeDate.BackColor = &H8000000F
'      'textFeeDate.Enabled = False
'      'textFeeDate.TabStop = False
'   End If
   
   ' 緬甸公報
   If m_TM10 = "048" Then
      EnableTextBox TextMyanmar_1, True
      EnableTextBox TextMyanmar_2, True
      EnableTextBox TextMyanmar_3, True
   Else
      EnableTextBox TextMyanmar_1, False
      EnableTextBox TextMyanmar_2, False
      EnableTextBox TextMyanmar_3, False
   End If
   
   Set rsTmp = Nothing
   
   '910729 Sieg
   m_TM21 = ""
   'm_TM22 = ""  '2006/12/26 CANCEL BY SONIA 因為菲律賓延展後使用宣誓要用原專用期限止日算
   New_TM22 = ""
   
   Dim strKey(0 To 4) As String, strTmp As String
   strKey(0) = m_CP09
   strKey(1) = m_TM01
   strKey(2) = m_TM02
   strKey(3) = m_TM03
   strKey(4) = m_TM04
   '910917 nick 應該要像內商一樣用商標種類區分
   '***** start
   'If TFGetMoneyDate(m_TM10, strKey, m_TM21, strTmp, m_TM22) Then
   'End If
   If m_TM01 = "CFT" Then
        'add by nick 2004/09/16 檢查是否有TG
         frm03010303_04.Hide
        Set frm03010303_04.UpForm = Me
        frm03010303_04.TGKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
        frm03010303_04.AllClass = textTM09 'edit by nickc 2006 " "
        frm03010303_04.Hide
        frm03010303_04.QueryData
        Unload frm03010303_04
        'edit by nick 2004/10/05
        'If ChkTG = True Then
            cmdInputTG.BackColor = &H8000000F
         'edit by nick 2004/10/05
         'Else
         '   cmdInputTG.BackColor = &HFF&
         'End If
      '91.9.19 modify by sonia
      'Select Case m_TM08
      'Case "1", "4", "7", "8":
      '      If TFGetMoneyDate(m_TM10, strKey, m_TM21, strTmp, m_TM22) Then
      '         If m_TM22 <> "0" Then m_TM22 = CompDate(2, 1, m_TM22)
      '      End If
      'Case Else
      '      strExc(0) = "SELECT TM22 FROM TRADEMARK WHERE TM15 = '" & m_TM27 & "' "
      '      intI = 1
      '      Set rsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
      '      If intI = 1 Then
      '         If Not IsNull(rsTemp.Fields("TM22")) Then
      '            m_TM22 = TransDate(rsTemp.Fields("TM22"), 1)
      '         End If
      '      End If
      'End Select
      '91.9.19 end
      '91.11.10 MODIFY BY SONIA
      'add by nick 2004/12/16 因為會殘留上一筆的值
      NickTmNa12 = 0
      Select Case m_CP10
      'modify by sonia 2025/9/4 +109緩審延展CFT-016520
      Case "102", "109":
            'modify by sonia 2014/10/31 加CP53延展前專用期止日CFT-14866莫三比克延展後五年要提使用宣誓
            If CFTGetNewDate(m_TM10, strKey, m_TM21, New_TM22, Pre_TM22) Then
            End If
      Case Else
            If TFGetMoneyDate(m_TM10, strKey, m_TM21, strTmp, New_TM22) Then
            End If
            'add by sonia 2022/8/31延展已發文才發註冊證案件改為延展後專用期止日CFT-014012
            'modify by sonia 2025/9/4 +109緩審延展CFT-016520
            strExc(0) = "select max(cp54) cp54 from caseprogress where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp27>19221111 and cp10 in ('102','109') "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If Val("" & RsTemp.Fields("cp54")) > 0 Then New_TM22 = "" & RsTemp.Fields("cp54")
            End If
            'end 2022/8/31
            'Added by Lydia 2025/09/02 奈及利亞商標管制緩審延展期限：相關總收文為申請101且該案曾發文緩審延展109則專用期止日的預設改為國家檔之商標專用年度NA13+延展年度NA14
            bolAddNa14 = False
            If m_TM01 = "CFT" And m_CP10 = "101" And m_TM10 = "302" Then
                strExc(0) = "select cp158 from caseprogress where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp27>19221111 and cp10='109' "
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                If intI = 1 Then
                    'New_TM22 = CompDate(0, m_NA14, New_TM22)  'cancel by sonia 2025/9/4 改在上面那句加109緩審延展，也許已經多次緩審延展
                    bolAddNa14 = True
                End If
            End If
            'end 2025/09/02
      End Select
      '91.11.10 END
   End If
   '***** end
End Sub

'edit by nick
'Public sub OnSaveData()
Public Function OnSaveData() As Boolean
OnSaveData = True
   Dim nIndex As Integer
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strCP06 As String
   Dim strCP07 As String
   Dim strCP10 As String
   'Dim strCP12 As String
   Dim strCP20 As String
   Dim strCP27 As String
   Dim strCP32 As String
   Dim strNP07 As String
   Dim strNP08 As String
   Dim strNP09 As String
   Dim strNP14 As String
   Dim strNP15 As String
   Dim strNP22 As String
   Dim strNA38 As String
   Dim bInsert As Boolean
   Dim strTemp As String
   'Add By Cheng 2002/06/07
   Dim strNA78 As String
   'Add By Sindy 2009/06/04
   Dim str202CP43 As String, str202CP10 As String, str202CP09 As String
   
 '911106 nick transation
On Error GoTo CheckingErr

cnnConnection.BeginTrans
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 案件性質為延展時
   ' 1. 更新原收文資料的實際結果為准, 延展註冊號數為輸入的審定號數
   ' 案件性質非延展時
   ' 1. 更新商標基本檔的發證日為來函收文日,專用期間, 目前准駁為准, 專用權是否存在為Y
   ' 2. 更新服務業務基本檔的發證日為來函收文日, 專用期間, 登記號, 註冊號數
   '案件性質為"延展"(102)時
   'modify by sonia 2025/9/4 +109緩審延展CFT-016520
   If m_CP10 = "102" Or m_CP10 = "109" Then
      'Modify By Cheng 2002/07/11
      '93.10.20 MODIFY BY SONIA
      'If m_TM10 = "019" Then
'2013/9/10 modify by sonia 泰國及阿根廷改同其他國家
'      If m_TM10 = "019" Or m_TM10 = "118" Then
'      '93.10.20 END
'         strSql = "UPDATE CaseProgress SET CP24 = '" & "1" & "', " & _
'                                          "CP30 = '" & textTM15 & "' " & _
'                  "WHERE CP09 = '" & m_CP09 & "' "
'         cnnConnection.Execute strSql
'      'add by nickc 2006/12/21
'      ElseIf m_TM10 = "046" Then
      If m_TM10 = "046" Then
'2013/9/10 end
         '2013/9/10 modify by sonia 應記錄原註冊號數
         'strSql = "UPDATE CaseProgress SET CP24 = '" & "1" & "', " & _
                                          "CP25=" & DBDATE(Text2) & ", " & _
                                          "CP30 = '" & textTM15 & "' " & _
                  "WHERE CP09 = '" & m_CP09 & "' "
         strSql = "UPDATE CaseProgress SET CP24 = '" & "1" & "', " & _
                                          "CP25=" & DBDATE(Text2) & ", " & _
                                          "CP30 = '" & IIf(textTM15S = textTM15, "", textTM15S) & "', " & _
                                          "CP64 = '" & IIf(textTM15S = textTM15, "", "原註冊號數為" & textTM15S & ";") & "'||CP64 " & _
                  "WHERE CP09 = '" & m_CP09 & "' "
         '2013/9/10 end
         cnnConnection.Execute strSql
         strSql = "UPDATE TradeMark SET TM15 = " & CNULL(Me.textTM15.Text) & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "'"
         cnnConnection.Execute strSql    '2013/9/10 add by sonia 發現少了這一句,所以基本檔都沒更新
      Else
         '2013/9/10 modify by sonia 應記錄原註冊號數
         'strSql = "UPDATE CaseProgress SET CP24 = '" & "1" & "', " & _
                                          "CP30 = '" & textTM15 & "' " & _
                  "WHERE CP09 = '" & m_CP09 & "' "
         strSql = "UPDATE CaseProgress SET CP24 = '" & "1" & "', " & _
                                          "CP30 = '" & IIf(textTM15S = textTM15, "", textTM15S) & "', " & _
                                          "CP64 = '" & IIf(textTM15S = textTM15, "", "原註冊號數為" & textTM15S & ";") & "'||CP64 " & _
                  "WHERE CP09 = '" & m_CP09 & "' "
         '2013/9/10 end
         cnnConnection.Execute strSql
         strSql = "UPDATE TradeMark SET TM15 = " & CNULL(Me.textTM15.Text) & _
                  " WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "'"
         cnnConnection.Execute strSql    '2013/9/10 add by sonia 發現少了這一句,所以基本檔都沒更新
      End If
      Select Case m_TM01
         Case "CFT":
            '93.10.28 MODIFY BY SONIA 加入下面之更新
            'strSQL = "UPDATE TradeMark SET TM16 = '1', " & _
            '                              "TM17 = 'Y', " & _
            '                              "TM21 = " & DBDATE(textTM21) & ", " & _
            '                              "TM22 = " & DBDATE(textTM22) & " " & _
            '         "WHERE TM01 = '" & m_TM01 & "' AND " & _
            '               "TM02 = '" & m_TM02 & "' AND " & _
            '               "TM03 = '" & m_TM03 & "' AND " & _
            '               "TM04 = '" & m_TM04 & "'"
            strSql = "UPDATE TradeMark SET TM14 = " & CNULL(Me.textTM14.Text) & ", " & _
                                          "TM16 = '1', " & _
                                          "TM17 = 'Y', " & _
                                          "TM21 = " & DBDATE(textTM21) & ", " & _
                                          "TM22 = " & DBDATE(textTM22) & " " & _
                     "WHERE TM01 = '" & m_TM01 & "' AND " & _
                           "TM02 = '" & m_TM02 & "' AND " & _
                           "TM03 = '" & m_TM03 & "' AND " & _
                           "TM04 = '" & m_TM04 & "'"
            cnnConnection.Execute strSql
            '93.10.28 END
         Case Else:
            '93.10.28 MODIFY BY SONIA 加入下面之更新
            'strSQL = "UPDATE ServicePractice SET SP20 = " & DBDATE(textTM21) & ", " & _
            '                                    "SP21 = " & DBDATE(textTM22) & " " & _
            '         "WHERE SP01 = '" & m_TM01 & "' AND " & _
            '               "SP02 = '" & m_TM02 & "' AND " & _
            '               "SP03 = '" & m_TM03 & "' AND " & _
            '               "SP04 = '" & m_TM04 & "' "
            strSql = "UPDATE ServicePractice SET SP14 = '" & textTM15 & "' " & _
                                                "SP20 = " & DBDATE(textTM21) & ", " & _
                                                "SP21 = " & DBDATE(textTM22) & " " & _
                     "WHERE SP01 = '" & m_TM01 & "' AND " & _
                           "SP02 = '" & m_TM02 & "' AND " & _
                           "SP03 = '" & m_TM03 & "' AND " & _
                           "SP04 = '" & m_TM04 & "' "
            '93.10.28 END
            cnnConnection.Execute strSql
      End Select
   '案件性質非"延展"(102)時
   Else
      Select Case m_TM01
         Case "CFT":
            'Modify By Cheng 2002/06/07
            '發文日以輸入的註冊日存入
'            strSQL = "UPDATE TradeMark SET TM16 = '1', " & _
'                                          "TM17 = 'Y', " & _
'                                          "TM20 = " & DBDATE(m_CP05) & ", " & _
'                                          "TM21 = " & DBDATE(textTM21) & ", " & _
'                                          "TM22 = " & DBDATE(textTM22) & " " & _
'                     "WHERE TM01 = '" & m_TM01 & "' AND " & _
'                           "TM02 = '" & m_TM02 & "' AND " & _
'                           "TM03 = '" & m_TM03 & "' AND " & _
'                           "TM04 = '" & m_TM04 & "'"
            'Modify By Cheng 2002/07/11
'            strSQL = "UPDATE TradeMark SET TM16 = '1', " & _
'                                          "TM17 = 'Y', " & _
'                                          "TM20 = " & DBDATE(Me.textRegDate.Text) & ", " & _
'                                          "TM21 = " & DBDATE(textTM21) & ", " & _
'                                          "TM22 = " & DBDATE(textTM22) & " " & _
'                     "WHERE TM01 = '" & m_TM01 & "' AND " & _
'                           "TM02 = '" & m_TM02 & "' AND " & _
'                           "TM03 = '" & m_TM03 & "' AND " & _
'                           "TM04 = '" & m_TM04 & "'"
            '93.10.28 MODIFY BY SONIA 加入下面之更新
            'strSQL = "UPDATE TradeMark SET TM16 = '1', " & _
            '                              "TM17 = 'Y', " & _
            '                              "TM20 = " & CNULL(Me.textRegDate.Text) & ", " & _
            '                              "TM21 = " & DBDATE(textTM21) & ", " & _
            '                              "TM22 = " & DBDATE(textTM22) & " " & _
            '         "WHERE TM01 = '" & m_TM01 & "' AND " & _
            '               "TM02 = '" & m_TM02 & "' AND " & _
            '               "TM03 = '" & m_TM03 & "' AND " & _
            '               "TM04 = '" & m_TM04 & "'"
            strSql = "UPDATE TradeMark SET TM14 = " & CNULL(Me.textTM14.Text) & ", " & _
                                          "TM15 = " & CNULL(Me.textTM15.Text) & ", " & _
                                          "TM16 = '1', " & _
                                          "TM17 = 'Y', " & _
                                          "TM20 = " & CNULL(Me.textRegDate.Text) & ", " & _
                                          "TM21 = " & DBDATE(textTM21) & ", " & _
                                          "TM22 = " & DBDATE(textTM22) & " " & _
                     "WHERE TM01 = '" & m_TM01 & "' AND " & _
                           "TM02 = '" & m_TM02 & "' AND " & _
                           "TM03 = '" & m_TM03 & "' AND " & _
                           "TM04 = '" & m_TM04 & "'"
            cnnConnection.Execute strSql
            '93.10.28 END
            'Add By Cheng 2002/07/22
            '若原商標基本檔准駁通知日(TM13)為NULL
            If m_TM13 = "" Then
               strSql = "UPDATE TradeMark SET TM13 = " & CNULL(Me.textRegDate.Text) & " " & _
                        " WHERE TM01 = '" & m_TM01 & "' AND " & _
                               "TM02 = '" & m_TM02 & "' AND " & _
                               "TM03 = '" & m_TM03 & "' AND " & _
                               "TM04 = '" & m_TM04 & "'"
               cnnConnection.Execute strSql
            End If
            
            'Add By Sindy 2009/06/04
            '抓其案號101.申請 或 701.領証且無CP24
            'Modify By Sindy 2015/2/26 再加入 107跨類, 308分割 案件性質
            '                          因為抓四個案件性質的資料可能同時存在一筆以上,請改以迴圈方式,
            '                          逐筆去更新CP24, CP25及下一程序的催審期限 (CFT-14892)
            'Modify By Sindy 2019/10/16 + 302奈及利亞國家增加檢查109緩審延展的催審
            'modify by sonia 2025/9/4 CFT-016520取消109緩審延展，只能在延展證書才能更新109緩審延展的催審
            strSql = "SELECT CP09,CP10 FROM CaseProgress WHERE CP01='" & m_TM01 & "' AND CP02='" & m_TM02 & "' AND CP03='" & m_TM03 & "' AND CP04='" & m_TM04 & "'" & _
                                                 " AND CP10 in ('101','701','107','308') AND CP24 is null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               RsTemp.MoveFirst
               Do While Not RsTemp.EOF
                  'cancel by sonia 2025/9/4 延展證書才能更新109緩審延展的催審CFT-016520
                  ''Modify By Sindy 2019/10/16 + 109緩審延展的催審
                  'If RsTemp("CP10") = "109" Then
                  '      strSql = "UPDATE NextProgress SET NP06 = '" & "N" & "' " & _
                  '                     "WHERE NP01 = '" & RsTemp("CP09") & "' AND " & _
                  '                           "NP02 = '" & m_TM01 & "' AND " & _
                  '                           "NP03 = '" & m_TM02 & "' AND " & _
                  '                           "NP04 = '" & m_TM03 & "' AND " & _
                  '                           "NP05 = '" & m_TM04 & "' AND " & _
                  '                           "NP06 IS NULL AND NP07 IN (305) "
                  '      cnnConnection.Execute strSql
                  '   End If  'Added by Lydia 2025/09/02
                  'Else
                  '2019/10/16 END
                     '該筆收文資料若無CP24時,才須更新CP24,CP25
                     strSql = "UPDATE CaseProgress SET CP24 = '1', " & _
                                                      "CP25=" & CNULL(DBDATE(textRegDate)) & " " & _
                                    "WHERE CP09 = '" & RsTemp("CP09") & "' "
                     cnnConnection.Execute strSql
                  'End If   'cancel by sonia 2025/9/4
                  '更新下一程序催審期限為Y
                  '2011/7/13 MODIFY BY SONIA 加入NP06條件且同時更新收達及提申 CFT-013505
                  strSql = "UPDATE NextProgress SET NP06 = '" & "Y" & "' " & _
                                 "WHERE NP01 = '" & RsTemp("CP09") & "' AND " & _
                                       "NP02 = '" & m_TM01 & "' AND " & _
                                       "NP03 = '" & m_TM02 & "' AND " & _
                                       "NP04 = '" & m_TM03 & "' AND " & _
                                       "NP05 = '" & m_TM04 & "' AND " & _
                                       "NP06 IS NULL AND NP07 IN (305,997,998) "
                  cnnConnection.Execute strSql
                  
                  RsTemp.MoveNext
               Loop
            '2015/2/26 END
            End If
            
            '案件性質為202.答辯且無CP24的收文資料
            strSql = "SELECT * FROM CaseProgress WHERE CP01='" & m_TM01 & "' AND CP02='" & m_TM02 & "' AND CP03='" & m_TM03 & "' AND CP04='" & m_TM04 & "' AND CP10='202' AND CP24 is null "
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
               With rsTmp
                  rsTmp.MoveFirst
                  Do While Not rsTmp.EOF
                     str202CP43 = "" & rsTmp.Fields("CP43")
                     str202CP10 = "" & rsTmp.Fields("CP10")
                     str202CP09 = "" & rsTmp.Fields("CP09")
                     '以CP43.相關總收文號一直往前串資料
                     Do While str202CP43 <> ""
                        strSql = "SELECT * FROM CaseProgress WHERE CP09 = '" & str202CP43 & "' "
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                        If intI = 1 Then
                           str202CP43 = "" & RsTemp.Fields("CP43")
                           str202CP10 = "" & RsTemp.Fields("CP10")
                        End If
                     Loop
                     '最後串到的案件性質為申請時
                     '才須更新此筆答辯的CP24,CP25及下一程序催審期限
                     If str202CP10 = "101" Then
                        strSql = "UPDATE CaseProgress SET CP24 = '1', " & _
                                                            "CP25=" & DBDATE(textRegDate) & " " & _
                                       "WHERE CP09 = '" & str202CP09 & "' "
                        cnnConnection.Execute strSql
                        strSql = "UPDATE NextProgress SET NP06 = '" & "Y" & "' " & _
                                 "WHERE NP01 = '" & str202CP09 & "' AND " & _
                                       "NP02 = '" & m_TM01 & "' AND " & _
                                       "NP03 = '" & m_TM02 & "' AND " & _
                                       "NP04 = '" & m_TM03 & "' AND " & _
                                       "NP05 = '" & m_TM04 & "' AND " & _
                                       "NP07 = " & "305"
                        cnnConnection.Execute strSql
                     End If
                     rsTmp.MoveNext
                  Loop
               End With
            End If
            rsTmp.Close
            Set rsTmp = Nothing
            '2009/06/04 End
            
            'Modify By Sindy 2019/10/16 + 302奈及利亞國家增加檢查109緩審延展
            strSql = "UPDATE NextProgress SET NP06 = '" & "N" & "' " & _
                     "WHERE NP02 = '" & m_TM01 & "' AND " & _
                           "NP03 = '" & m_TM02 & "' AND " & _
                           "NP04 = '" & m_TM03 & "' AND " & _
                           "NP05 = '" & m_TM04 & "' AND " & _
                           "(NP06 IS NULL) AND NP07 IN ('109') "
            cnnConnection.Execute strSql
            '2019/10/16 END
            
            'Add By Sindy 2013/6/3
            'CFT美國案,更新下一程序檔,案件性質為通知使用宣誓的期限資料
            If m_TM10 = "101" Then
               strSql = "UPDATE NextProgress SET NP06 = '" & "Y" & "' " & _
                        "WHERE NP02 = '" & m_TM01 & "' AND " & _
                              "NP03 = '" & m_TM02 & "' AND " & _
                              "NP04 = '" & m_TM03 & "' AND " & _
                              "NP05 = '" & m_TM04 & "' AND " & _
                              "(NP06 IS NULL) AND NP07 IN ('1711') "
               cnnConnection.Execute strSql
            End If
            '2013/6/3 END
            
         Case Else:
            'Modify By Cheng 2002/06/07
            '發文日以輸入的註冊日存入
'            strSQL = "UPDATE ServicePractice SET SP12 = " & DBDATE(m_CP05) & ", " & _
'                                                "SP20 = " & DBDATE(textTM21) & ", " & _
'                                                "SP21 = " & DBDATE(textTM22) & " " & _
'                     "WHERE SP01 = '" & m_TM01 & "' AND " & _
'                           "SP02 = '" & m_TM02 & "' AND " & _
'                           "SP03 = '" & m_TM03 & "' AND " & _
'                           "SP04 = '" & m_TM04 & "' "
            'Modify By Cheng 2002/07/11
'            strSQL = "UPDATE ServicePractice SET SP12 = " & DBDATE(Me.textRegDate.Text) & ", " & _
'                                                "SP20 = " & DBDATE(textTM21) & ", " & _
'                                                "SP21 = " & DBDATE(textTM22) & " " & _
'                     "WHERE SP01 = '" & m_TM01 & "' AND " & _
'                           "SP02 = '" & m_TM02 & "' AND " & _
'                           "SP03 = '" & m_TM03 & "' AND " & _
'                           "SP04 = '" & m_TM04 & "' "
            '93.10.28 MODIFY BY SONIA 加入下面之更新
            'strSQL = "UPDATE ServicePractice SET SP12 = " & CNULL(Me.textRegDate.Text) & ", " & _
            '                                    "SP20 = " & DBDATE(textTM21) & ", " & _
            '                                    "SP21 = " & DBDATE(textTM22) & " " & _
            '         "WHERE SP01 = '" & m_TM01 & "' AND " & _
            '               "SP02 = '" & m_TM02 & "' AND " & _
            '               "SP03 = '" & m_TM03 & "' AND " & _
            '               "SP04 = '" & m_TM04 & "' "
            strSql = "UPDATE ServicePractice SET SP14 = '" & textTM15 & "', " & _
                                                "SP12 = " & CNULL(Me.textRegDate.Text) & ", " & _
                                                "SP20 = " & CNULL(DBDATE(textTM21)) & ", " & _
                                                "SP21 = " & CNULL(DBDATE(textTM22)) & " " & _
                     "WHERE SP01 = '" & m_TM01 & "' AND " & _
                           "SP02 = '" & m_TM02 & "' AND " & _
                           "SP03 = '" & m_TM03 & "' AND " & _
                           "SP04 = '" & m_TM04 & "' "
            '93.10.28 END
            cnnConnection.Execute strSql
            ' 系統類別為CFC時需更新登記號及註冊號數
            If m_TM01 = "CFC" Then
               strSql = "UPDATE ServicePractice SET SP13 = " & CNULL(textSP13) & ", " & _
                                                   "SP14 = '" & textTM15 & "' " & _
                        "WHERE SP01 = '" & m_TM01 & "' AND " & _
                              "SP02 = '" & m_TM02 & "' AND " & _
                              "SP03 = '" & m_TM03 & "' AND " & _
                              "SP04 = '" & m_TM04 & "' "
               cnnConnection.Execute strSql
               
               'Add By Sindy 2015/2/26 另非CFT案件(服務業務之CFC)時, 也抓該案號的806.著作權登記且無CP24者
               '                       更新CP24 , CP25及下一程序的催審期限 (CFC-000757)
               strSql = "SELECT CP09 FROM CaseProgress WHERE CP01='" & m_TM01 & "' AND CP02='" & m_TM02 & "' AND CP03='" & m_TM03 & "' AND CP04='" & m_TM04 & "'" & _
                                                       " AND CP10 in ('806') AND CP24 is null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  RsTemp.MoveFirst
                  Do While Not RsTemp.EOF
                     '該筆收文資料若無CP24時,才須更新CP24,CP25
                     strSql = "UPDATE CaseProgress SET CP24 = '1', " & _
                                                      "CP25=" & CNULL(DBDATE(textRegDate)) & " " & _
                                    "WHERE CP09 = '" & RsTemp("CP09") & "' "
                     cnnConnection.Execute strSql
                     '更新下一程序催審期限為Y
                     strSql = "UPDATE NextProgress SET NP06 = '" & "Y" & "' " & _
                                    "WHERE NP01 = '" & RsTemp("CP09") & "' AND " & _
                                          "NP02 = '" & m_TM01 & "' AND " & _
                                          "NP03 = '" & m_TM02 & "' AND " & _
                                          "NP04 = '" & m_TM03 & "' AND " & _
                                          "NP05 = '" & m_TM04 & "' AND " & _
                                          "NP06 IS NULL AND NP07 IN (305,997,998) "
                     cnnConnection.Execute strSql
                     RsTemp.MoveNext
                  Loop
               End If
               '2015/2/26 END
            End If
      End Select
   End If
   'Add By Cheng 2002/07/11
   '93.10.28 CANCEL BY SONIA 移至上面
   'Select Case m_TM01
   '   Case "CFT"
   '      strSQL = "UPDATE TradeMark SET TM14 = " & CNULL(Me.textTM14.Text) & ", " & _
   '                                    "TM15 = " & CNULL(Me.textTM15.Text) & "  " & _
   '               "WHERE TM01 = '" & m_TM01 & "' AND " & _
   '                     "TM02 = '" & m_TM02 & "' AND " & _
   '                     "TM03 = '" & m_TM03 & "' AND " & _
   '                     "TM04 = '" & m_TM04 & "'"
   '   Case Else
   '      strSQL = "UPDATE ServicePractice SET SP14 = '" & textTM15 & "' " & _
   '               "WHERE SP01 = '" & m_TM01 & "' AND " & _
   '                     "SP02 = '" & m_TM02 & "' AND " & _
   '                     "SP03 = '" & m_TM03 & "' AND " & _
   '                     "SP04 = '" & m_TM04 & "' "
   'End Select
   'cnnConnection.Execute strSQL
   '93.10.28 END
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 以本所案號+來函收文日+案件性質(註冊證)來檢查案件進度檔, 若不存在則新增一筆
   bInsert = True
   ' 收文號
   strCP09 = Empty
   ' SQL 語法
   '2006/9/29 MODIFY BY SONIA 應分1701及1713
   'strSQL = "SELECT * FROM CaseProgress " & _
   '         "WHERE CP01 = '" & m_TM01 & "' AND " & _
   '               "CP02 = '" & m_TM02 & "' AND " & _
   '               "CP03 = '" & m_TM03 & "' AND " & _
   '               "CP04 = '" & m_TM04 & "' AND " & _
   '               "CP05 = " & DBDATE(m_CP05) & " AND " & _
   '               "CP10 = '" & "1701" & "' "
   'modify by sonia 2025/9/4 +109緩審延展CFT-016520
   If m_CP10 = "102" Or m_CP10 = "109" Then
      strSql = "SELECT * FROM CaseProgress " & _
               "WHERE CP01 = '" & m_TM01 & "' AND " & _
                     "CP02 = '" & m_TM02 & "' AND " & _
                     "CP03 = '" & m_TM03 & "' AND " & _
                     "CP04 = '" & m_TM04 & "' AND " & _
                     "CP05 = " & DBDATE(m_CP05) & " AND " & _
                     "CP10 = '" & "1713" & "' "
   Else
      strSql = "SELECT * FROM CaseProgress " & _
               "WHERE CP01 = '" & m_TM01 & "' AND " & _
                     "CP02 = '" & m_TM02 & "' AND " & _
                     "CP03 = '" & m_TM03 & "' AND " & _
                     "CP04 = '" & m_TM04 & "' AND " & _
                     "CP05 = " & DBDATE(m_CP05) & " AND " & _
                     "CP10 = '" & "1701" & "' "
   End If
   '2006/9/29 END
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      bInsert = False
      rsTmp.MoveFirst
      ' 若原先已有資料存在則使用原有的收文號
      strCP09 = rsTmp.Fields("CP09")
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   ' 若要新增資料到案件進度檔則需產生新的收文號
   If bInsert = True Then
      strCP09 = AutoNo("C", 6)
   End If
   
   ' 有領證費用時, 是否向客戶收款及是否開電腦收據為空白, 否則為N
   strCP20 = Empty
   strCP32 = Empty
   If IsEmptyText(textFee) = True Then
      strCP20 = "N"
      strCP32 = "N"
   End If
   ' 案件性質為註冊證
   '92.10.2 MODIFY BY SONIA
   'StrCp10 = "1701"
   'modify by sonia 2025/9/4 +109緩審延展CFT-016520
   If m_CP10 = "102" Or m_CP10 = "109" Then
      strCP10 = "1713"
   Else
      strCP10 = "1701"
   End If
   ' 業務區別 91.8.26 MODIFY BY SONIA
   'strCP12 = GetStaffDepartment(m_CP13)
   ' 發文日
   strCP27 = DBDATE(SystemDate())
   ' 新增案件進度資料或更新
   If bInsert = True Then
      'Modify By Cheng 2002/07/11
      '若案件性質為"1701"時, 其CP43應存frm03010303_02所點選的收文號
'      strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32) " & _
'               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                       "'" & strCP09 & "','" & strCP10 & "','" & strCP12 & "','" & m_CP13 & "','" & strUserNum & "'," & _
'                       "'" & strCP20 & "','" & "N" & "'," & strCP27 & ",'" & strCP32 & "') "
        'Modify By Cheng 2003/04/03
        '智權人員存最近收文A類接洽記錄單的智權人員
        'Modify By Cheng 2004/02/04
        '業務區為最近收文A類接洽記錄單智權人員的業務區
'      strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43) " & _
'               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                       "'" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "'," & _
'                       "'" & strCP20 & "','" & "N" & "'," & strCP27 & ",'" & strCP32 & "'," & CNULL(m_CP09) & ") "
      strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                       "'" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "'," & _
                       "'" & strCP20 & "','" & "N" & "'," & strCP27 & ",'" & strCP32 & "'," & CNULL(m_CP09) & ") "
        'End
      cnnConnection.Execute strSql
      
        'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
        Pub_UpdateFromMaxCP27 m_TM01, m_TM02, m_TM03, m_TM04
      
   Else
      'Modify By Sindy 2010/10/20 cp12,cp13抓最新的智權人員資料
      strSql = "UPDATE CaseProgress SET CP05 = " & DBDATE(m_CP05) & ", " & _
                                       "CP12 = '" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "', " & _
                                       "CP13 = '" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "', " & _
                                       "CP14 = '" & strUserNum & "', " & _
                                       "CP20 = '" & strCP20 & "', " & _
                                       "CP26 = '" & "N" & "', " & _
                                       "CP27 = " & strCP27 & ", " & _
                                       "CP32 = '" & strCP32 & "', " & _
                                       "CP43 = '" & m_CP09 & "' " & _
               "WHERE CP09 = '" & strCP09 & "' "
      Pub_SeekTbLog strSql 'Add By Sindy 2020/10/30 CFT-017832 第1次有輸報價,但又做第2次沒輸領證費
      cnnConnection.Execute strSql
   End If
   ' 費用
   If Val(textFee) > 0 Then
      'StrSQL = "UPDATE CASEPROGRESS SET CP18 = " & textFee & " " & _
      '         "WHERE CP09 = '" & strCP09 & "' "
      strSql = "UPDATE CASEPROGRESS SET CP16 = " & textFee & " " & _
               "WHERE CP09 = '" & strCP09 & "' "
      'Add By Sindy 2020/10/30 CFT-017832 第1次有輸報價,但又做第2次沒輸領證費
      If bInsert = False Then
         Pub_SeekTbLog strSql
      End If
      '2020/10/30 END
      cnnConnection.Execute strSql
   Else
      'StrSQL = "UPDATE CASEPROGRESS SET CP18 = " & "NULL" & " " & _
      '         "WHERE CP09 = '" & strCP09 & "' "
      strSql = "UPDATE CASEPROGRESS SET CP16 = " & "NULL" & " " & _
               "WHERE CP09 = '" & strCP09 & "' "
      'Add By Sindy 2020/10/30 CFT-017832 第1次有輸報價,但又做第2次沒輸領證費
      If bInsert = False Then
         Pub_SeekTbLog strSql
      End If
      '2020/10/30 END
      cnnConnection.Execute strSql
   End If
   ' 點數
   If Val(textCP18) > 0 Then
      'StrSQL = "UPDATE CASEPROGRESS SET CP20 = " & textCP20 & " " & _
      '         "WHERE CP09 = '" & strCP09 & "' "
      strSql = "UPDATE CASEPROGRESS SET CP18 = " & textCP18 & " " & _
               "WHERE CP09 = '" & strCP09 & "' "
      'Add By Sindy 2020/10/30 CFT-017832 第1次有輸報價,但又做第2次沒輸領證費
      If bInsert = False Then
         Pub_SeekTbLog strSql
      End If
      '2020/10/30 END
      cnnConnection.Execute strSql
   Else
      'StrSQL = "UPDATE CASEPROGRESS SET CP20 = " & "NULL" & " " & _
      '         "WHERE CP09 = '" & strCP09 & "' "
      strSql = "UPDATE CASEPROGRESS SET CP18 = " & "NULL" & " " & _
               "WHERE CP09 = '" & strCP09 & "' "
      'Add By Sindy 2020/10/30 CFT-017832 第1次有輸報價,但又做第2次沒輸領證費
      If bInsert = False Then
         Pub_SeekTbLog strSql
      End If
      '2020/10/30 END
      cnnConnection.Execute strSql
   End If
   If Val(textFee) > 0 Then
      'Add By Cheng 2002/08/23
      strSql = "UPDATE CASEPROGRESS SET CP17 = CP16 - (nvl(CP18,0)*1000) " & _
               "WHERE CP09 = '" & strCP09 & "' "
      'Add By Sindy 2020/10/30 CFT-017832 第1次有輸報價,但又做第2次沒輸領證費
      If bInsert = False Then
         Pub_SeekTbLog strSql
      End If
      '2020/10/30 END
      cnnConnection.Execute strSql
   End If
      
   '2006/3/21 ADD BY SONIA
   ' 更新下一程序檔案件性質為催審的資料
   '2011/7/13 MODIFY BY SONIA 從下面併進來997,998及NP06條件
   strSql = "UPDATE NextProgress SET NP06 = '" & "Y" & "' " & _
            "WHERE NP01 = '" & m_CP09 & "' AND " & _
                  "NP02 = '" & m_TM01 & "' AND " & _
                  "NP03 = '" & m_TM02 & "' AND " & _
                  "NP04 = '" & m_TM03 & "' AND " & _
                  "NP05 = '" & m_TM04 & "' AND " & _
                  "(NP06 IS NULL OR NP06 <> 'Y') AND NP07 IN (305,997,998) "
   cnnConnection.Execute strSql
   '2006/3/21 END
      
   'Add By Sindy 2009/06/22
   '在註冊証前提申之使用宣誓，輸入註冊証時，一併將使用宣誓之催審期限上Y
   If m_TM01 = "CFT" And Me.textRegDate.Text <> "" Then
      strSql = "UPDATE NextProgress SET NP06 = 'Y' " & _
            "where (np01,np07,np02,np03,np04,np05,np22) in ( " & _
               "SELECT np01,np07,np02,np03,np04,np05,np22 " & _
               "From caseprogress, NextProgress " & _
               "Where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' " & _
               "and cp10='105' " & _
               "and cp27 < " & DBDATE(textRegDate) & " " & _
               "and cp27 is not null and cp27<>0 " & _
               "and cp09=np01(+) " & _
               "and cp01=np02(+) and cp02=np03(+) and cp03=np04(+) and cp04=np05(+) " & _
               "and np07(+)='305') "
      cnnConnection.Execute strSql
   End If
   '2009/06/22 End
   
   '2009/1/19 MODIFY BY SONIA 加CFC不控制,CFC-000769
   If m_TM01 <> "CFC" Then
      bInsert = True   'ADD BY SONIA 2021/9/15
      ' 新增延展記錄到下一程序檔
       'Modify By Cheng 2003/06/12
       '若有新增進度檔才抓新的下一程序流水號
   '   strNP22 = GetNextProgressNo()
      '2009/10/29 ADD BY SONIA 因延展期限可能於延展已提申時先掛下一程序,故此重新判斷是否已掛
      strSql = "SELECT * FROM NextProgress " & _
               "WHERE NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND " & _
                     "NP06 IS NULL AND " & _
                     "NP07 = '102' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         m_NP22 = rsTmp.Fields("NP22") 'Add By Sindy 2011/5/24 若原先已有資料存在則取得原有的序號
         bInsert = False
      End If
      rsTmp.Close
      '2009/10/29 END
      
      strNP07 = "102"
      ' 法定期限為專用期限之截止日
      strNP09 = DBDATE(textTM22)
      'Modify By Cheng 2002/06/07
   '   ' 本所期限為法定期限-2天
   '   strNp08 = DBDATE(DateSerial(Val(DBYEAR(strNp09)), Val(DBMONTH(strNp09)), Val(DBDAY(strNp09)) - 2))
      ' 本所期限為法定期限-以申請國家抓國家檔之延展時間(月)
       'Modify By Cheng 2003/09/02
   '   strNP08 = DBDATE(DateSerial(Val(DBYEAR(strNP09)), Val(DBMONTH(strNP09)) - GetDelayTime(m_TM10), Val(DBDAY(strNP09))))
      '2005/9/28 MODIFY BY SONIA 業務要求改為提前2個月
      'strNP08 = DBDATE(DateAdd("m", -GetDelayTime(m_TM10), ChangeWStringToWDateString(DBDATE(strNP09))))
      ' 本所期限為法定期限-2個月
      strNP08 = CompDate(1, -2, strNP09)
      strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      m_NP08 = strNP08
      ' 組成SQL語法
       '若為新增進度檔
       If bInsert = True Then
           'Modify By Cheng 2003/04/03
           '智權人員存最近收文A類接洽記錄單的智權人員
           strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                    "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                              strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & GetNextProgressNo() & ")"
       '若為更新進度檔
       Else
           '2009/10/29 MODIFY BY SONIA 以本所案號更新
           'strSQL = "Update NextProgress Set NP08=" & strNP08 & ", NP09=" & strNP09 & " " & _
                           " Where  NP01='" & strCP09 & "' And NP07='102' And NP06 Is Null "
'           strSql = "Update NextProgress Set NP08=" & strNP08 & ", NP09=" & strNP09 & " " & _
'                    " Where NP02 = '" & m_TM01 & "' And NP03 = '" & m_TM02 & "' And NP04 = '" & m_TM03 & "' And NP05 = '" & m_TM04 & "' And NP07='102' And NP06 Is Null "
'           '2009/10/29 MODIFY BY SONIA 若為延展證書同時更新總收文號為C類收文號
'           If strCP10 = "1713" Then
'              cnnConnection.Execute strSql
              'Modify By Sindy 2015/3/10 更新原下一程序延展期限時,NP01同時也要更新為CP09,資料才會一致
              strSql = "Update NextProgress Set NP01='" & strCP09 & "',NP08=" & strNP08 & ",NP09=" & strNP09 & _
                       " Where NP02 = '" & m_TM01 & "' And NP03 = '" & m_TM02 & "' And NP04 = '" & m_TM03 & "' And NP05 = '" & m_TM04 & "' And NP07='102' And NP06 Is Null"
'           End If
           '2009/10/29 END
       End If
      cnnConnection.Execute strSql
   End If
   
   'Add By Sindy 2023/5/3
   '新增內部收文
   If Me.Text3.Text <> "" Then
       'Modified by Lydia 2024/04/17 更改註冊證=> IIf(m_CP10 = "102", "更改延展證書", "更改註冊證")
       'modify by sonia 2025/9/4 +109緩審延展也是更改延展證書CFT-016520
       strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP26,CP32,CP43,CP64,CP20) " & _
                       "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strSrvDate(1) & "," & _
                       "'" & AutoNo("B", 6) & "','302','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
                       "'" & strUserNum & "','','N','" & strCP09 & "', '" & IIf(m_CP10 = "102", "更改延展證書", IIf(m_CP10 = "109", "更改延展證書", "更改註冊證")) & "','N')"
       cnnConnection.Execute strSql
   End If
   '2023/5/3 END
    
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'Modify By Cheng 2002/06/07
   '改為若有輸入繳年費期限時, 只新增第1筆資料到下一程序檔
'   ' 若有輸入繳年費期限時, 新增2筆資料到下一程序檔
'   ' 1. 第一筆的繳年費期限為輸入的繳年費期限
'   ' 2. 第二筆的繳年費期限為專用期限止日加十年
   If IsEmptyText(textFeeDate) = False Then
      ' 第一筆
      ' 案件性質為繳年費
      strNP07 = "708"
      ' 法定期限為繳年費期限
      strNP09 = DBDATE(textFeeDate)
      'Modify By Cheng 2002/06/07
'      ' 本所期限為法定期限-2天
'      strNp08 = DBDATE(DateSerial(Val(DBYEAR(strNp09)), Val(DBMONTH(strNp09)), Val(DBDAY(strNp09)) - 2))
      ' 本所期限為法定期限 - 1年
        'Modify By Cheng 2003/09/02
'      strNP08 = DBDATE(DateSerial(Val(DBYEAR(strNP09)) - 1, Val(DBMONTH(strNP09)), Val(DBDAY(strNP09))))
      strNP08 = DBDATE(DateAdd("yyyy", -1, ChangeWStringToWDateString(DBDATE(strNP09))))
      strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      ' 組成SQL語法
        'Modify By Cheng 2003/04/03
        '智權人員存最近收文A類接洽記錄單的智權人員
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                          strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & GetNextProgressNo() & ")"
      cnnConnection.Execute strSql
      
      'Modify By Cheng 2002/06/07
      '取消新增資料
'      ' 第二筆
'      ' 案件性質為繳年費
'      strNP07 = "708"
'      ' 法定期限為繳年費期限
'      strTemp = DBDATE(textTM21)
'      strNp09 = DBDATE(DateSerial(Val(DBYEAR(strTemp)) + 5, Val(DBMONTH(strTemp)), Val(DBDAY(strTemp))))
'      ' 本所期限為法定期限-2天
'      strNp08 = DBDATE(DateSerial(Val(DBYEAR(strNp09)), Val(DBMONTH(strNp09)), Val(DBDAY(strNp09)) - 2))
'      ' 組成SQL語法
'      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'                "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
'                          strNp08 & "," & strNp09 & ",'" & m_CP13 & "'," & GetNextProgressNo() & ")"
'      cnnConnection.Execute strSQL
   End If
   
   
   'add by nickc  2006/12/21 柬埔寨使用宣誓期間歸零
   m_046Date2 = ""
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 申請國家為美國, 菲律賓, 柬埔寨, 案件性質為申請或延展時, 新增資料到下一程序檔
   '2005/3/23 modify by sonia
   'If (m_TM10 = "101" Or m_TM10 = "030" Or m_TM10 = "046") And (m_CP10 = "101" Or m_CP10 = "102") Then
   'edit by nickc 2007/02/15 加入葡萄牙
   'If (m_TM10 = "101" Or m_TM10 = "030" Or m_TM10 = "046") And (m_CP10 = "101" Or m_CP10 = "102" Or m_CP10 = "308") Then
   '2008/10/13 modify by sonia 取消葡萄牙
   'If (m_TM10 = "101" Or m_TM10 = "030" Or m_TM10 = "046" Or m_TM10 = "213") And (m_CP10 = "101" Or m_CP10 = "102" Or m_CP10 = "308") Then
   'modify by sonia 2014/10/31 加入莫三比克318
   'Modify By Sindy 2015/2/26 不限定申請國家,依國家設定為主
   'If (m_TM10 = "101" Or m_TM10 = "030" Or m_TM10 = "046" Or m_TM10 = "318") And (m_CP10 = "101" Or m_CP10 = "102" Or m_CP10 = "308") Then
   'modify by sonia 2025/9/4 +109緩審延展CFT-016520
   If m_CP10 = "101" Or m_CP10 = "102" Or m_CP10 = "308" Or m_CP10 = "109" Then
      'Add/Modify By Cheng 2002/06/07
      Select Case m_CP10
      Case "101", "308" '申請,分割
         ' 取得使用宣誓年度
         strNA38 = 0
         If rsTmp.State <> adStateClosed Then rsTmp.Close
         Set rsTmp = Nothing
         strSql = "SELECT * FROM Nation WHERE NA01 = '" & m_TM10 & "' AND NA38 IS NOT NULL "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            strNA38 = rsTmp.Fields("NA38")
            ' 案件性質為使用宣誓
            strNP07 = "105"
            'Modify By Cheng 2002/06/07
            ' 法定期限為專用期起日 + 使用宣誓年度
            '2011/9/23 全部改以註冊日計算
            'strTemp = DBDATE(Me.textTM21.Text)
            strTemp = DBDATE(textRegDate)
            '2011/9/23 END
            'modify by sonia 2022/10/22 +318莫三比克也是申請日起算CFT-020729
            'If m_TM10 = "112" Then strTemp = DBDATE(m_TM11)   'add by sonia 2017/12/14  波多黎各112改用申請日CFT-014266
            If m_TM10 = "112" Or m_TM10 = "318" Then
               strTemp = DBDATE(m_TM11)
            End If
            'end 2022/10/21
            strNP09 = DBDATE(DateAdd("yyyy", Val(strNA38), ChangeWStringToWDateString(DBDATE(strTemp))))
            'add by sonia 2018/11/16   '墨西哥核准日期(即發證日或註冊日)落在2018/8/10當天或之後者，管制三年使用宣誓期限，即註冊日起滿三年後之三個月內應提出使用宣誓
            'modify by sonia 2023/9/15 海地110法定期限要再加3個月CFT-023278
            If m_TM10 = "104" Or m_TM10 = "110" Then
               strNP09 = CompDate(1, 3, strNP09)
            End If
            'end  2018/11/16
            ' 本所期限為法定期限-2天
            '若申請國家為"菲律賓"時, 本所期限 = 法定期限 - 半年, 其他國家則 本所期限 = 法定期限 - 1年
            'edit by  nickc 2007/05/01 業務說改成本所=法定-2個月 不管任何國家
            strNP08 = CompDate(1, -2, strNP09)
            strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
            '92.10.22 end
         '若無使用宣誓年度則不新增下一程序檔
         Else
            rsTmp.Close
            Set rsTmp = Nothing
            GoTo NextLine
         End If
         rsTmp.Close
         Set rsTmp = Nothing
      'modify by sonia 2025/9/4 +109緩審延展CFT-016520
      'Case "102" '延展
      Case Else '102延展,109緩審延展
         'modify by sonia 2015/12/4 改為延展後使用宣誓年度NA78,原來抓下次使用宣誓年度NA39
         strNA78 = 0
         If rsTmp.State <> adStateClosed Then rsTmp.Close
         Set rsTmp = Nothing
         strSql = "SELECT * FROM Nation WHERE NA01 = '" & m_TM10 & "' AND NA78 IS NOT NULL "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            strNA78 = rsTmp.Fields("NA78")
            ' 案件性質為使用宣誓
            strNP07 = "105"
            ' 法定期限為該案件原來之專用期止日 + 下次使用宣誓年度
            'edit by nickc 2006/12/21 柬埔寨以畫面延展核准日計算
            If m_TM10 = "046" Then
               strTemp = DBDATE(Text2)
            Else
               'modify by sonia 2014/10/31 加CP53延展前專用期止日CFT-14866莫三比克延展後五年要提使用宣誓
               'strTemp = DBDATE(m_TM22)
               strTemp = DBDATE(Pre_TM22)
               'end 2014/10/31
            End If
            'Modify By Cheng 2003/09/02
'            strNP09 = DBDATE(DateSerial(Val(DBYEAR(strTemp)) + Val(strNA78), Val(DBMONTH(strTemp)), Val(DBDAY(strTemp))))
            strNP09 = DBDATE(DateAdd("yyyy", Val(strNA78), ChangeWStringToWDateString(DBDATE(strTemp))))
            If m_TM10 = "110" Then strNP09 = CompDate(1, 3, Val(strNP09))   'add by sonia 2023/9/15 海地110法定期限要再加3個月CFT-023278
'2011/9/23 modify by sonia  2007/05/01 業務說改成本所=法定-2個月 不管任何國家, 但此處沒改到
''            ' 本所期限為法定期限-2天
'            '若申請國家為"菲律賓"時, 本所期限 = 法定期限 - 半年, 其他國家則 本所期限 = 法定期限 - 1年
            strNP08 = CompDate(1, -2, strNP09)
            strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
'2011/9/23 end
            'add by nickc 2006/12/21
            If m_TM10 = "046" Then
                m_046Date2 = strNP09
            End If
         '若無下次使用宣誓年度則不新增下一程序檔
         Else
            rsTmp.Close
            Set rsTmp = Nothing
            GoTo NextLine
         End If
         rsTmp.Close
         Set rsTmp = Nothing
      End Select
      
      'Add By Sindy 2009/11/04 判斷是否已掛使用宣誓期限
      bInsert = True
      strSql = "SELECT * FROM NextProgress " & _
               "WHERE NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND " & _
                     "NP06 IS NULL AND " & _
                     "NP07 = '" & strNP07 & "' "
      '2010/9/20因菲律賓申請日+3年為第一次使用宣誓期限,以前發證都會在此之前故可直接更新下一程序檔使用宣誓期限
      '         但近日發現菲律賓已提早發證故不可直接更新,否則會把第一次期限蓋掉CFT-12549
      '         改為只更新該註冊證產生的使用宣誓期限
      'modify by sonia 2014/10/31 僅菲律賓要加條件,其他國家不可
      'strSql = strSql & " and np01='" & strCP09 & "' "
      'modify by sonia 2015/12/10 且只能為註冊證,延展證書則不可,否則延展代理人提申時產生之期限不會更新
      'modify by sonia 2020/12/28 +波多黎各112
      'If m_TM10 = "030" And m_CP10 <> "102" Then strSql = strSql & " and np01='" & strCP09 & "' "
      If (m_TM10 = "030" Or m_TM10 = "112") And m_CP10 <> "102" Then strSql = strSql & " and np01='" & strCP09 & "' "
      'end 2014/10/31
      '2010/9/20 end
      'add by sonia 2022/6/10 菲律賓原延展期限後一年有掛使用宣誓期限,不可蓋掉,故加入NP09>Pre_TM22+20000原延展法定期限2年的條件
      If (m_TM10 = "030" Or m_TM10 = "112") And m_CP10 = "102" Then strSql = strSql & "and NP09>" & Pre_TM22 + 20000
      'end 2022/6/10
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         bInsert = False
      End If
      '2009/11/04 End
            
      ' 組成SQL語法
      'Modify By Cheng 2003/04/03
      '智權人員存最近收文A類接洽記錄單的智權人員
      'Modify By Sindy 2009/11/04
      If bInsert = False Then
         strSql = "Update NextProgress Set NP01='" & strCP09 & "',NP08=" & strNP08 & ", NP09=" & strNP09 & " " & _
                   "WHERE NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND " & _
                     "NP06 IS NULL AND " & _
                     "NP07 = '" & strNP07 & "' "
         'modify by sonia 2014/10/31 僅菲律賓要加條件,其他國家不可
         'strSql = strSql & " and np01='" & strCP09 & "' "   '2010/9/20 add by sonia
         'modify by sonia 2015/12/10 且只能為註冊證,延展證書則不可,否則延展代理人提申時產生之期限不會更新
         'modify by sonia 2020/12/28 +波多黎各112
         If (m_TM10 = "030" Or m_TM10 = "112") And m_CP10 <> "102" Then strSql = strSql & " and np01='" & strCP09 & "' "
         'end 2014/10/31
         'add by sonia 2022/6/10 菲律賓原延展期限後一年有掛使用宣誓期限,不可蓋掉,故加入NP09>Pre_TM22+20000原延展法定期限2年的條件
         If (m_TM10 = "030" Or m_TM10 = "112") And m_CP10 = "102" Then strSql = strSql & "and NP09>" & Pre_TM22 + 20000
         'end 2022/6/10
      
      '2009/11/04 End
      Else
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                   "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                             strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & GetNextProgressNo() & ")"
      End If
      cnnConnection.Execute strSql
      'add by sonia 2019/10/14 期限若已過期要提醒CFT-002361
      If DBDATE(strNP08) < Val(strSrvDate(1)) Then
         MsgBox "下次使用宣誓期限已過期, 請注意!!!", vbExclamation + vbOKOnly
      End If
      'end 2019/10/14
     
NextLine: '不新增下一程序檔
   
   End If
   Set rsTmp = Nothing
   
   'add by nickc 2005/04/22
   Pub_UpdateEndModCash m_TM01, m_TM02, m_TM03, m_TM04
   
   'Add by Sindy 2023/4/27
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm03010303_01", strCP09
   End If
   '2023/4/27 END
   
   '911106 nick transation
   cnnConnection.CommitTrans
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
'  Set rsTmp = Nothing
' '911106 nick transation
'  cnnConnection.CommitTrans
     Exit Function
CheckingErr:
    cnnConnection.RollbackTrans
    'Modified by Lydia 2019/11/21 CFT-18187輸入註冊證,一直彈出遺漏表示式(by A1028) ; 可是電腦中心人員測不到
    'MsgBox (Err.Description)
    If Err.Number <> 0 Then
         MsgBox Err.Description & vbCrLf & strSql
    'add by nick 2004/11/03
    End If
    OnSaveData = False
End Function

Private Sub Form_Unload(Cancel As Integer)
    'Add By Cheng 2002/07/19
   Set frm03010303_03 = Nothing
End Sub

'add by nickc 2006/12/21
Private Sub Text2_GotFocus()
InverseTextBox Text2
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(Text2) = False Then
      If CheckIsDate(Text2, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "延展核准日日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Text2_GotFocus
         Exit Sub
      End If
   End If
End Sub

'Add By Sindy 2023/5/3
Private Sub Text3_GotFocus()
    'Memo by Lydia 2024/04/17 標題從「是否更改證書」改為「是否更改註冊證/延展證書」
    TextInverse Me.Text3
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 89 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
'2023/5/3 END

' 點數
Private Sub textCP18_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCP18) = False Then
      If IsNumeric(textCP18) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "點數請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP18_GotFocus
      Else
         If IsEmptyText(textFee) = True Then
            Cancel = True
            strTit = "資料檢核"
            strMsg = "請同時輸入領證費及點數"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textFee_GotFocus
         End If
      End If
   Else
      If IsEmptyText(textFee) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請同時輸入領證費及點數"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP18_GotFocus
      End If
   End If
End Sub

' 繳年費期限
Private Sub textFeeDate_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textFeeDate) = False Then
      If CheckIsDate(textFeeDate, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "繳年費期限日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFeeDate_GotFocus
         Exit Sub
      End If
      'Add By Cheng 2002/03/11
      If Val(Me.textFeeDate.Text) < ServerDate Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "繳年費期限日期不可小於系統日期!!!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFeeDate_GotFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否列印定稿
Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textPrint) = False Then
      Select Case textPrint
         Case "", " ", "N":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrint_GotFocus
      End Select
   End If
End Sub

' 註冊日
Private Sub textRegDate_Validate(Cancel As Boolean)
   Dim SysDate As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textRegDate) = False Then
      ' 檢查是否為民國年
      If CheckIsDate(textRegDate, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的註冊日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textRegDate_GotFocus
      End If
      ' 註冊日不可超過系統日
      If Val(DBDATE(textRegDate)) > Val(DBDATE(SystemDate())) Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "註冊日不可超過系統日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textRegDate_GotFocus
      End If
   End If
   '910919 nick 檢查定義若是應該與註冊日做檢查，則註冊日不能空白
   If NickTmNa12 = 6 Then
        If Trim(textRegDate) = "" Then
            Cancel = True
            strTit = "資料檢核"
            strMsg = "註冊日不能空白！"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textRegDate_GotFocus
        End If
   End If
End Sub

' 公告日
Private Sub textTM14_Validate(Cancel As Boolean)
   Dim SysDate As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM14) = False Then
      ' 檢查是否為民國年
      If CheckIsDate(textTM14, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的公告日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM14_GotFocus
      End If
      ' 公告日不可超過系統日
      If Val(DBDATE(textTM14)) > Val(DBDATE(SystemDate())) Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "公告日不可超過系統日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM14_GotFocus
      End If
   End If
   '910919 nick 檢查定義若是應該與公告日做檢查，則公告日不能空白
   If NickTmNa12 = 5 Then
        If Trim(textTM14) = "" Then
            Cancel = True
            strTit = "資料檢核"
            strMsg = "公告日不能空白！"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM14_GotFocus
        End If
   End If
End Sub

' 2006/4/11 ADD BY SONIA 緬甸延展日報日期
Private Sub Text1_Validate(Cancel As Boolean)
   Dim SysDate As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(Text1) = False Then
      If CheckIsDate(Text1, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的緬甸延展日報日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Text1_GotFocus
      End If
      If Val(DBDATE(Text1)) > Val(DBDATE(SystemDate())) Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "緬甸延展日報日期不可超過系統日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Text1_GotFocus
      End If
   End If
   '緬甸延展案時緬甸延展日報日期不能空白,定稿會印
   If m_CP10 = "102" And m_TM10 = "048" Then
      If Trim(Text1) = "" Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "緬甸延展日報日期不能空白！"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Text1_GotFocus
      End If
   End If
End Sub
'2006/4/11 END
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
'add by nick 2004/09/17
'edit by nick 2004/10/05
'If ChkTG = False And cmdInputTG.Visible = True Then
'         strTit = "資料檢核"
'         strMsg = "商品及服務必須要輸入！"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         GoTo EXITSUB
'End If

   ' 申請國家為西班牙時
   '2005/3/23 modify by sonia
   'If m_TM10 = "211" And m_CP10 = "101" Then
   'Modify By Sindy 2012/9/20 西班牙已取消繳年費制度, 故請取消畫面 繳年費期限 的輸入及控制
'   If m_TM10 = "211" And (m_CP10 = "101" Or m_CP10 = "308") Then
'      If IsEmptyText(textFeeDate) = True Then
'         strTit = "資料檢核"
'         strMsg = "申請國家為西班牙, 繳年費期限一定要輸入"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textFeeDate.SetFocus
'         GoTo EXITSUB
'      End If
'   End If
   'add by nickc 2006/12/21
   If m_TM10 = "046" And m_CP10 = "102" Then
      If IsEmptyText(Text2) = True Then
         strTit = "資料檢核"
         strMsg = "申請國家為柬埔寨, 延展核准日一定要輸入"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Text2.SetFocus
         GoTo EXITSUB
      End If
   End If
   'Add By Cheng 2002/03/11
   If Me.textFeeDate.Text <> "" Then
      If Val(Me.textFeeDate.Text) < ServerDate Then
         strTit = "資料檢核"
         strMsg = "繳年費期限不可小於系統日期!!!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFeeDate.SetFocus
         GoTo EXITSUB
      End If
   End If
   '910919 nick 檢查定義若是應該與註冊日做檢查，則註冊日不能空白
   If NickTmNa12 = 6 Then
        If Trim(textRegDate) = "" Then
            strTit = "資料檢核"
            strMsg = "註冊日不能空白！"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textRegDate.SetFocus
            GoTo EXITSUB
        End If
   End If
   '910919 nick 檢查定義若是應該與公告日做檢查，則公告日不能空白
   If NickTmNa12 = 5 Then
        If Trim(textTM14) = "" Then
            strTit = "資料檢核"
            strMsg = "公告日不能空白！"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM14.SetFocus
            GoTo EXITSUB
        End If
   End If
   '910919 nick 專用期限不可空白
   '2009/1/19 MODIFY BY SONIA 加CFC不控制,CFC-000769
   If m_TM01 <> "CFC" Then
      If Trim(textTM21) = "" Or Trim(textTM22) = "" Then
           strTit = "資料檢核"
           strMsg = "專用期限不能空白！"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           textTM21.SetFocus
           GoTo EXITSUB
      End If
   End If
   
   'Add By Sindy 2013/3/14
   If Trim(textTM22) <> "" Then
      If Val(DBDATE(textTM22)) < Val(strSrvDate(1)) Then
         If MsgBox("確定專用期限的止日要小於系統日嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
            textTM22.SetFocus
            GoTo EXITSUB
         End If
      End If
   End If
   '2013/3/14 End
   
   'Add By Cheng 2004/03/16
   '若申請國家為巴拿馬, 則註冊日不可空白
   '93.11.2 MODIFY BY SONIA
   'If m_TM10 = "103" Then
   '2005/3/23 modify by sonia
   'If m_TM10 = "103" And (m_CP10 = "101" Or m_CP10 = "107") Then
   If m_TM10 = "103" And (m_CP10 = "101" Or m_CP10 = "107" Or m_CP10 = "308") Then
   '93.11.2 END
       If Me.textRegDate.Text = "" Then
           strTit = "資料檢核"
           strMsg = "註冊日不能空白！"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           Me.textRegDate.SetFocus
           GoTo EXITSUB
       End If
   End If
   'End
   
   'Added by Lydia 2020/11/09 若前一畫面傳進來的案件性質是申請或分割時，請加判斷若該案號沒有領證701進度時，此畫面上之領證費及點數一定要輸入，反之一定不能輸入。
   strExc(1) = ""
   strExc(0) = "select cp09 from caseprogress where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp57 is null and cp10='701' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
       strExc(1) = RsTemp.Fields("cp09")
   End If
   If m_CP10 = "101" Or m_CP10 = "308" Then
       'modify by sonia 2021/1/5 不必管點數CFT-021569
       'If strExc(1) = "" And (Val(textFee) = 0 Or Val(textCP18) = 0) Then
       If strExc(1) = "" And Val(textFee) = 0 Then
           MsgBox "本案無〔領證〕進度，請輸入領證費及點數！", vbExclamation, "檢核資料"
           If Val(textFee) = 0 Then
               textFee.SetFocus
               textFee_GotFocus
           Else
               textCP18.SetFocus
               textCP18_GotFocus
           End If
           GoTo EXITSUB
       End If
       If strExc(1) <> "" And (Val(textFee) > 0 Or Val(textCP18) > 0) Then
           MsgBox "本案有〔領證〕進度，請勿輸入領證費及點數！", vbExclamation, "檢核資料"
           textFee.Text = ""
           textCP18.Text = ""
       End If
   End If
   'end 2020/11/09
   
   'add by nickc 2007/04/17 加入若是輸入領証費時，檢查是否已收文過領証且未取消收文的，若是有就問是否要加收費用，選是就加收，選無就讓他清掉
   If textFee <> "" Then
       'Modified by Lydia 2020/11/09 改在前面判斷收領證費
       'CheckOC3
       'strSql = "select * from caseprogress where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp57 is null and cp10='701' "
       'AdoRecordSet3.CursorLocation = adUseClient
       'AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       'If AdoRecordSet3.RecordCount <> 0 Then
       If strExc(1) <> "" Then
       'end 2020/11/09
           If MsgBox("已收文領証費，是否要加收費用？", vbYesNo) = vbNo Then
               textFee.SetFocus
               GoTo EXITSUB
           End If
       End If
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textTM15_GotFocus()
   InverseTextBox textTM15
End Sub

'Add By Sindy 2010/9/1
Private Sub textTM15_Validate(Cancel As Boolean)
Dim strRetrunText As String 'Add By Sindy 2017/5/17
   
   If IsEmptyText(textTM15) = False Then
      '檢查審定號所輸入的長度是否正確
      'Add By Sindy 2017/5/17 + strRetrunText
      If PUB_ChkTm12Tm15Length("2", textTM15, m_TM01, m_TM02, m_TM03, m_TM04, m_TM10, , , strRetrunText) = False Then
         Cancel = True
         textTM15_GotFocus
         Exit Sub
      'Add By Sindy 2017/5/17
      Else
         textTM15 = strRetrunText
      '2017/5/17 END
      End If
   End If
End Sub

Private Sub textTM21_GotFocus()
   InverseTextBox textTM21
End Sub

' 專用期限起日
Private Sub textTM21_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM21) = False Then
      ' 檢查是否為西元日期
      If CheckIsDate(textTM21, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "專用期限起日日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo A0
      End If
      
      Dim strTmp As String
      '910917 nick   應該是 CFT
      '910919 再度修正當資料庫中沒值時要與輸入值做比對
      '***** start
      'If m_TM01 = "TF" Then
      'edit by nickc  2007/11/13 義大利延展不管專用期間
      'If m_TM01 = "CFT" Then
      If m_TM01 = "CFT" And Not (m_CP10 = "102" And m_TM10 = "204") Then
         'If Trim(m_TM21) <> "" Then
         '   If textTM21 <> m_TM21 Then
         '      Cancel = True
         '      strTit = "資料檢核"
         '      strMsg = "專用期限起日應為<" & m_TM21 & ">"
         '      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '      GoTo A0
         '   End If
         'Else
            Select Case NickTmNa12
            Case 5     '公告日
                 If textTM14 <> textTM21 Then
                    Cancel = True
                    strTit = "資料檢核"
                    strMsg = "專用期限起日應為<" & textTM14 & ">"
                    nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                    GoTo A0
                 End If
            Case 6     '發證日
                  'add by sonia 2022/1/19 墨西哥修法20201105以前申請案件為申請日起10年
                  If m_TM10 = "104" And m_TM11 < 20201105 Then
                     If textTM21 <> m_TM11 Then
                         Cancel = True
                         strTit = "資料檢核"
                         strMsg = "專用期限起日應為<" & m_TM11 & ">"
                         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                         GoTo A0
                     End If
                  Else
                  'end 2022/1/19
                     If textRegDate <> textTM21 Then
                         Cancel = True
                         strTit = "資料檢核"
                         strMsg = "專用期限起日應為<" & textRegDate & ">"
                         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                         GoTo A0
                     End If
                  End If    'add by sonia 2022/1/19
           Case Else
                '93.8.5 modify by sonia 智利不檢查專用期限
                'If textTM21 <> m_TM21 Then
                If textTM21 <> m_TM21 And m_TM10 <> "126" Then
                '93.8.5 end
                   Cancel = True
                   strTit = "資料檢核"
                   strMsg = "專用期限起日應為<" & m_TM21 & ">"
                   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                   GoTo A0
                End If
            End Select
         'End If
      End If
      '***** end
      ' 申請國家為西班牙時, 繳年費期限為專用期限起日加上5年
      '2005/3/23 modify by sonia
      'If m_TM10 = "211" And m_CP10 = "101" Then
      'Modify By Sindy 2012/9/20 西班牙已取消繳年費制度, 故請取消畫面 繳年費期限 的輸入及控制
'      If m_TM10 = "211" And (m_CP10 = "101" Or m_CP10 = "308") Then
'         If IsEmptyText(textTM21) = False Then
'            'Modify By Cheng 2003/09/02
''            textFeeDate = DBDATE(DateSerial(Val(DBYEAR(textTM21)) + 5, Val(DBMONTH(textTM21)), Val(DBDAY(textTM21))))
'            textFeeDate = DBDATE(DateAdd("yyyy", 5, ChangeWStringToWDateString(DBDATE(textTM21))))
'         End If
'      End If
   End If
A0:
   If Cancel Then TextInverse textTM21
   
End Sub

' 專用期限止日
Private Sub textTM22_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
'Add By Cheng 2003/03/03
Dim strReservedDate As String '專用期限
   
   Cancel = False
   If IsEmptyText(textTM22) = False Then
      ' 檢查是否為西元日期
      If CheckIsDate(textTM22, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "專用期限止日日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo A0
      End If
      
      Dim strTmp As String
      '910919 修正當資料庫中沒值時要與輸入值做比對
      '***** start
      'If m_TM01 = "CFT" Then
      '   If textTM22 <> m_TM22 Then
      '      Cancel = True
      '      strTit = "資料檢核"
      '      strMsg = "專用期限止日應為<" & m_TM22 & ">"
      '      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '   End If
      'End If
      'edit by nickc  2007/11/13 義大利延展不管專用期間
      'If m_TM01 = "CFT" Then
      If m_TM01 = "CFT" And Not (m_CP10 = "102" And m_TM10 = "204") Then
         'If Trim(m_TM22) <> "" Then
         '   If textTM22 <> m_TM22 Then
         '      Cancel = True
         '      strTit = "資料檢核"
         '      strMsg = "專用期限止日應為<" & m_TM22 & ">"
         '      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '      GoTo A0
         '   End If
         'Else
            Select Case NickTmNa12
            Case 5     '公告日
               If Val(Left(textTM14, 4)) + NickTmNa13 & Right(textTM14, 4) <> textTM22 And CompDate(2, -1, (Val(Left(textTM14, 4)) + NickTmNa13 & Right(textTM14, 4))) <> textTM22 Then
                  Cancel = True
                  strTit = "資料檢核"
                  strMsg = "專用期限止日應為<" & Val(Left(textTM14, 4)) + NickTmNa13 & Right(textTM14, 4) & "> 或 <" & CompDate(2, -1, (Val(Left(textTM14, 4)) + NickTmNa13 & Right(textTM14, 4))) & ">"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  GoTo A0
               End If
            Case 6     '發證日
                Select Case m_TM10
                'Add By Cheng 2003/04/01
                '菲律賓的商標專用期間可為10年或20年
'2009/8/24 modify by sonia 葉易雲說全部改為10年
'                Case "030"
'                    If (Val(Left(textRegDate, 4)) + 10 & Right(textRegDate, 4) <> textTM22 And CompDate(2, -1, (Val(Left(textRegDate, 4)) + 10 & Right(textRegDate, 4))) <> textTM22) And _
'                        Val(Left(textRegDate, 4)) + 20 & Right(textRegDate, 4) <> textTM22 And CompDate(2, -1, (Val(Left(textRegDate, 4)) + 20 & Right(textRegDate, 4))) <> textTM22 Then
'                       Cancel = True
'                       strTit = "資料檢核"
'                       strMsg = "專用期限止日應為<" & Val(Left(textRegDate, 4)) + 10 & Right(textRegDate, 4) & "> 或 <" & CompDate(2, -1, (Val(Left(textRegDate, 4)) + 10 & Right(textRegDate, 4))) & ">" & vbCrLf & _
'                                    "　　　　　　　或<" & Val(Left(textRegDate, 4)) + 20 & Right(textRegDate, 4) & "> 或 <" & CompDate(2, -1, (Val(Left(textRegDate, 4)) + 20 & Right(textRegDate, 4))) & ">"
'                       nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'                       GoTo A0
'                    End If
'2009/8/24 end
                'Add By Sindy 2012/11/27
                Case "113" '委內瑞拉
                     '註冊日期在2008年9月17日之前者,專用期限自註冊日起算10年
                     '註冊日期在2008年9月17日之後者,專用期限自註冊日起算15年(國家檔設定)
                     If Val(DBDATE(textRegDate)) < 20080917 Then
                        If Val(Left(textRegDate, 4)) + 10 & Right(textRegDate, 4) <> textTM22 And CompDate(2, -1, (Val(Left(textRegDate, 4)) + 10 & Right(textRegDate, 4))) <> textTM22 Then
                           Cancel = True
                           strTit = "資料檢核"
                           strMsg = "專用期限止日應為<" & Val(Left(textRegDate, 4)) + 10 & Right(textRegDate, 4) & "> 或 <" & CompDate(2, -1, (Val(Left(textRegDate, 4)) + 10 & Right(textRegDate, 4))) & ">"
                           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                           GoTo A0
                        End If
                     Else
                        If Val(Left(textRegDate, 4)) + NickTmNa13 & Right(textRegDate, 4) <> textTM22 And CompDate(2, -1, (Val(Left(textRegDate, 4)) + NickTmNa13 & Right(textRegDate, 4))) <> textTM22 Then
                           Cancel = True
                           strTit = "資料檢核"
                           strMsg = "專用期限止日應為<" & Val(Left(textRegDate, 4)) + NickTmNa13 & Right(textRegDate, 4) & "> 或 <" & CompDate(2, -1, (Val(Left(textRegDate, 4)) + NickTmNa13 & Right(textRegDate, 4))) & ">"
                           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                           GoTo A0
                        End If
                     End If
                '2012/11/27 End
                'add by sonia 2019/6/19
                Case "102"  '加拿大
                     '發證日在2019/06/17以前者維持15年,之後改為10年；每次延展一律改為10年(原為15年)
                     If Val(DBDATE(textRegDate)) <= 20190617 Then
                        If Val(Left(textRegDate, 4)) + 15 & Right(textRegDate, 4) <> textTM22 And CompDate(2, -1, (Val(Left(textRegDate, 4)) + 15 & Right(textRegDate, 4))) <> textTM22 Then
                           Cancel = True
                           strTit = "資料檢核"
                           strMsg = "專用期限止日應為<" & Val(Left(textRegDate, 4)) + 15 & Right(textRegDate, 4) & "> 或 <" & CompDate(2, -1, (Val(Left(textRegDate, 4)) + 15 & Right(textRegDate, 4))) & ">"
                           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                           GoTo A0
                        End If
                     Else
                        If Val(Left(textRegDate, 4)) + NickTmNa13 & Right(textRegDate, 4) <> textTM22 And CompDate(2, -1, (Val(Left(textRegDate, 4)) + NickTmNa13 & Right(textRegDate, 4))) <> textTM22 Then
                           Cancel = True
                           strTit = "資料檢核"
                           strMsg = "專用期限止日應為<" & Val(Left(textRegDate, 4)) + NickTmNa13 & Right(textRegDate, 4) & "> 或 <" & CompDate(2, -1, (Val(Left(textRegDate, 4)) + NickTmNa13 & Right(textRegDate, 4))) & ">"
                           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                           GoTo A0
                        End If
                     End If
                'end 2019/6/19
                'add by sonia 2022/1/19
                Case "104"  '加拿大
                     '墨西哥修法20201105以前申請案件為申請日起10年
                     If m_TM11 < 20201105 Then
                        If Val(Left(m_TM11, 4)) + 10 & Right(m_TM11, 4) <> textTM22 And CompDate(2, -1, (Val(Left(m_TM11, 4)) + 10 & Right(m_TM11, 4))) <> textTM22 Then
                           Cancel = True
                           strTit = "資料檢核"
                           strMsg = "專用期限止日應為<" & Val(Left(m_TM11, 4)) + 10 & Right(m_TM11, 4) & "> 或 <" & CompDate(2, -1, (Val(Left(m_TM11, 4)) + 10 & Right(m_TM11, 4))) & ">"
                           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                           GoTo A0
                        End If
                     Else
                        If Val(Left(textRegDate, 4)) + NickTmNa13 & Right(textRegDate, 4) <> textTM22 And CompDate(2, -1, (Val(Left(textRegDate, 4)) + NickTmNa13 & Right(textRegDate, 4))) <> textTM22 Then
                           Cancel = True
                           strTit = "資料檢核"
                           strMsg = "專用期限止日應為<" & Val(Left(textRegDate, 4)) + NickTmNa13 & Right(textRegDate, 4) & "> 或 <" & CompDate(2, -1, (Val(Left(textRegDate, 4)) + NickTmNa13 & Right(textRegDate, 4))) & ">"
                           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                           GoTo A0
                        End If
                     End If
                'end 2022/1/19
                '其他國家
                Case Else
                    If Val(Left(textRegDate, 4)) + NickTmNa13 & Right(textRegDate, 4) <> textTM22 And CompDate(2, -1, (Val(Left(textRegDate, 4)) + NickTmNa13 & Right(textRegDate, 4))) <> textTM22 Then
                       Cancel = True
                       strTit = "資料檢核"
                       strMsg = "專用期限止日應為<" & Val(Left(textRegDate, 4)) + NickTmNa13 & Right(textRegDate, 4) & "> 或 <" & CompDate(2, -1, (Val(Left(textRegDate, 4)) + NickTmNa13 & Right(textRegDate, 4))) & ">"
                       nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                       GoTo A0
                    End If
                End Select
            Case Else
                'Modify By Cheng 2003/03/03
                Select Case m_TM10
                     'CANCEL BY SONIA 2019/3/5
'                     Case "231"    '德國
'                         If textTM22 <> IIf(New_TM22 <> "", ChangeWDateStringToWString(DateSerial(Mid(New_TM22, 1, 4), Mid(New_TM22, 5, 2), PUB_GetMonthDays(Mid(New_TM22, 1, 4), Mid(New_TM22, 5, 2)))), "") Then
'                             Cancel = True
'                             strTit = "資料檢核"
'                             strMsg = "專用期限止日應為<" & IIf(New_TM22 <> "", ChangeWDateStringToWString(DateSerial(Mid(New_TM22, 1, 4), Mid(New_TM22, 5, 2), PUB_GetMonthDays(Mid(New_TM22, 1, 4), Mid(New_TM22, 5, 2)))), "") & ">"
'                             nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'                             GoTo A0
'                         End If
                     '2005/5/23 ADD BY SONIA
                     Case "040"    '印度申請日19990101以前為7年, 國家檔設定10年
                         strReservedDate = New_TM22
                         'edit by nickc 2005/12/06 may 說改成 cft-006020 才可以過
                         'If m_TM11 < 19990101 Then m_CP10
                         '2006/3/8 MODIFY BY SONIA MAY說現在所有延展皆為10年
                         'If m_TM11 < 19970401 Then
                         'edit by nickc 2007/12/14 may 說不是很確定印度的時間，先再往前調一個月 ext: cft-008121
                         'If m_TM11 < 19970401 And m_CP10 = "101" Then
                         If m_TM11 < 19970301 And m_CP10 = "101" Then
                            strReservedDate = CompDate(0, -3, New_TM22)
                         End If
                         If textTM22 <> strReservedDate And textTM22 <> CompDate(2, -1, strReservedDate) Then
                            strTit = "資料檢核"
                            strMsg = "專用期限止日應為<" & strReservedDate & "> 或 <" & CompDate(2, -1, strReservedDate) & ">"
                            Cancel = True
                            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                            GoTo A0
                         End If
                     '2005/5/23 END
                     '92.11.8 ADD BY SONIA
                     '2008/11/24 MODIFY BY SONIA 加入巴基斯坦038
                     '2009/3/27 MODIFY BY SONIA 葉易雲說2004年4月12日前期滿應延展案件為15年,之後為10年,故取消巴基斯坦038
                     Case "018"   '馬來西亞 專用期為10年或7年 國家檔設定10年
                         If textTM22 <> New_TM22 And textTM22 <> CompDate(0, -3, New_TM22) Then
                             strTit = "資料檢核"
                             strMsg = "專用期限止日應為<" & New_TM22 & "> 或 <" & CompDate(0, -3, New_TM22) & ">"
                             Cancel = True
                             nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                             GoTo A0
                         End If
                     '93.8.5 add by sonia 智利126不檢查專用期限
                     '93.10.20 MODIFT BY SONIA 阿根廷118不檢查專用期限止日
                     Case "126", "118"
                     '93.8.5 end
                     '92.11.8 END
                     'cancel by sonia 2025/5/13 緬甸已修法改申請日起算，陳蒲璇說延展也要加入檢查，故取消CFT-22178
                     ''2006/4/11 ADD BY SONIA 依葉易雲需求,緬甸延展專用期止日不檢查
                     'Case "048"
                     '    If m_CP10 <> "102" Then
                     '        If Val(Left(textRegDate, 4)) + NickTmNa13 & Right(textRegDate, 4) <> textTM22 And CompDate(2, -1, (Val(Left(textRegDate, 4)) + NickTmNa13 & Right(textRegDate, 4))) <> textTM22 Then
                     '           Cancel = True
                     '           strTit = "資料檢核"
                     '           strMsg = "專用期限止日應為<" & Val(Left(textRegDate, 4)) + NickTmNa13 & Right(textRegDate, 4) & "> 或 <" & CompDate(2, -1, (Val(Left(textRegDate, 4)) + NickTmNa13 & Right(textRegDate, 4))) & ">"
                     '           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                     '           GoTo A0
                     '        End If
                     '    End If
                     ''2006/4/11 END
                     'end 2025/5/13
                     '其他國家
                     Case Else
                         If textTM22 <> New_TM22 And textTM22 <> CompDate(2, -1, New_TM22) Then
                             strTit = "資料檢核"
                             strMsg = "專用期限止日應為<" & New_TM22 & "> 或 <" & CompDate(2, -1, New_TM22) & ">"
                             '92.9.4 MODIFY BY SONIA
                             'Cancel = True
                             'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                             '若申請國家為沙烏地阿拉伯
                             'modify by sonia 2020/12/17 葉芳如說德國也要提醒但仍可以繼續輸入CFT-013877
                              If m_TM10 = "021" Or m_TM10 = "231" Then
                                  nResponse = MsgBox(strMsg & " ，已確認無誤繼續輸入嗎？", vbYesNo, strTit)
                                  If nResponse = vbNo Then Cancel = True
                              Else
                                   Cancel = True
                                   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                              End If
                              '92.9.4 END
                              GoTo A0
                         End If
                End Select
            End Select
         'End If
      End If
      '***** end
   End If
A0:
   If Cancel Then TextInverse textTM22
End Sub

Private Sub textTM22_GotFocus()
   InverseTextBox textTM22
End Sub

Private Sub textFeeDate_GotFocus()
   InverseTextBox textFeeDate
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textTM14_GotFocus()
   InverseTextBox textTM14
End Sub

Private Sub textSP13_GotFocus()
   InverseTextBox textSP13
End Sub

Private Sub textCP18_GotFocus()
   InverseTextBox textCP18
End Sub

Private Sub textRegDate_GotFocus()
   InverseTextBox textRegDate
End Sub

Private Sub textFee_GotFocus()
   InverseTextBox textFee
End Sub

Private Sub Text1_GotFocus()
   InverseTextBox Text1
End Sub

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
   Dim strSql As String
   Dim strTemp1 As String
   Dim strTemp2 As String
   Dim strLC02 As String 'Added by Morgan 2015/6/16 報價定稿用,取代原 m_NP22
   
   'Modified by Morgan 2015/6/16 自動發證的領證報價統一設 LC02=0
   'strLC02 = m_NP22
   strLC02 = "0"
   'end 2015/6/16
   strET03 = "" 'Add By Sindy 2023/5/3
   Select Case m_TM01
      Case "CFC":
         If IsCustomerIndividual(m_TM23) = True Then
            ' 清除定稿例外欄位檔原有資料
            EndLetter strET01, strCP09, "21", strUserNum
         Else
            If m_SP48 = "Y" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter strET01, strCP09, "19", strUserNum
            Else
               ' 清除定稿例外欄位檔原有資料
               EndLetter strET01, strCP09, "20", strUserNum
            End If
         End If
      Case "CFT":
         'add by nick 2004/10/05
         If ChkTG = True Then
            EndLetter strET01, strCP09, "26", strUserNum
         End If

         Select Case m_CP10
'--------102延展  ---memo by Lydia 2024/04/17
            ' 延展
            Case "102":
                'Modify By Cheng 2004/03/10
                Select Case m_TM10
                Case "121" '烏拉圭
                    ' 清除定稿例外欄位檔原有資料
                    strET03 = "19" 'Added by Lydia 2024/04/17
                    EndLetter strET01, strCP09, "19", strUserNum
                    '2013/10/11 ADD BY SONIA 原審定號數
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & strET01 & "','" & strCP09 & "','" & "19" & "','" & strUserNum & _
                             "','附註','" & textTM15S & "')"
                    cnnConnection.Execute strSql
                    '2013/10/11 END
                Case "206" '奧地利
                    ' 清除定稿例外欄位檔原有資料
                    strET03 = "20" 'Added by Lydia 2024/04/17
                    EndLetter strET01, strCP09, "20", strUserNum
                '2005/5/23 ADD BY SONIA 印度
                Case "040"
                    '2006/3/8 MODIFY BY SONIA MAY說取消此控制,延展全部十年
                    'If m_TM11 < 20010511 Then
                    '    ' 清除定稿例外欄位檔原有資料
                    '    EndLetter strET01, strCP09, "27", strUserNum
                    'Else
                    '    ' 清除定稿例外欄位檔原有資料
                    '    EndLetter strET01, strCP09, "18", strUserNum
                    'End If
                    ' 清除定稿例外欄位檔原有資料
                    strET03 = "18" 'Added by Lydia 2024/04/17
                    EndLetter strET01, strCP09, "18", strUserNum
                    '2006/3/8 END
               '2005/5/23 END
               '2006/4/11 ADD BY SONIA
                Case "048" '緬甸
                    ' 清除定稿例外欄位檔原有資料
                    'Modified by Lydia 2024/04/17 更正延展證書>> 改為一般
                    'EndLetter strET01, strCP09, "28", strUserNum
                    ' 法定期限
                    'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & strET01 & "','" & strCP09 & "','" & "28" & "','" & strUserNum & _
                             "','其他日期','" & DBDATE(Text1) & "')"
                    'cnnConnection.Execute strSql
                    If textTM15S = textTM15 Then
                        strTemp1 = "18"
                    Else
                        strTemp1 = "17"
                    End If
                    strET03 = strTemp1
                    EndLetter strET01, strCP09, strTemp1, strUserNum
                    ' 法定期限
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & strET01 & "','" & strCP09 & "','" & strTemp1 & "','" & strUserNum & _
                             "','其他日期','" & DBDATE(Text1) & "')"
                    cnnConnection.Execute strSql
                    'end 2024/04/17
               '2006/4/11 END
                'add by nickc 2006/11/13 柬埔寨
                Case "046"
                    strET03 = "29" 'Added by Lydia 2024/04/17
                    EndLetter strET01, strCP09, "29", strUserNum
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & strET01 & "','" & strCP09 & "','" & "29" & "','" & strUserNum & _
                           "','法定期限','" & DBDATE(m_046Date2) & "')"
                     cnnConnection.Execute strSql
                'add  by nickc 2007/03/03 加入紐西蘭
                Case "016"
                    strET03 = "30" 'Added by Lydia 2024/04/17
                    EndLetter strET01, strCP09, "30", strUserNum
                'add by nickc 2007/06/15  加入澳洲
                Case "015"
                    strET03 = "31" 'Added by Lydia 2024/04/17
                    EndLetter strET01, strCP09, "31", strUserNum
                'add by nickc 2007/07/10 加入韓國
                Case "012"
                    strET03 = "32" 'Added by Lydia 2024/04/17
                    EndLetter strET01, strCP09, "32", strUserNum
                'add by nickc 2007/08/28 加入波蘭
                Case "222"
                    strET03 = "33" 'Added by Lydia 2024/04/17
                    EndLetter strET01, strCP09, "33", strUserNum
                'add by nickc 2008/05/02
                '2009/3/10 MODIFY BY SONIA 取消委內瑞拉113
                Case "120", "115", "114", "116" '玻利維亞、厄瓜多、哥倫比亞、祕魯
                    ' 清除定稿例外欄位檔原有資料
                    '2013/10/11 MODIFY BY SONIA 判斷延展審定號數與原審定號數是否相同
                    'EndLetter strET01, strCP09, "34", strUserNum
                    If textTM15S = textTM15 Then
                        strET03 = "34" 'Added by Lydia 2024/04/17
                        EndLetter strET01, strCP09, "34", strUserNum
                    Else
                        strET03 = "24" 'Added by Lydia 2024/04/17
                        EndLetter strET01, strCP09, "24", strUserNum
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & "24" & "','" & strUserNum & _
                                 "','附註','" & textTM15S & "')"
                        cnnConnection.Execute strSql
                    End If
                    '2013/10/11 END
                    
                'Add By Sindy 2009/04/29
                Case "030" '菲律賓
                    ' 清除定稿例外欄位檔原有資料
                    strET03 = "35" 'Added by Lydia 2024/04/17
                    EndLetter strET01, strCP09, "35", strUserNum
                    'add by sonia 2023/2/14
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & strET01 & "','" & strCP09 & "','" & "35" & "','" & strUserNum & _
                             "','其他日期','" & Pre_TM22 & "')"
                    cnnConnection.Execute strSql
                    'end 2023/2/14
                '2009/04/29 End
                
                'Add By Sindy 2010/4/1
                Case "019" '泰國
                    ' 清除定稿例外欄位檔原有資料
                    strET03 = "36" 'Added by Lydia 2024/04/17
                    EndLetter strET01, strCP09, "36", strUserNum
                '2010/4/1 End
                
                'Modify By Sindy 2010/12/31 印尼與其他國家只差證書2字, 加拿大也要將核准通知改為證書2字, 因此直接改其他國家定稿, 印尼也不須另立定稿
'                'Add By Sindy 2010/11/17
'                Case "017": '印尼
'                    ' 清除定稿例外欄位檔原有資料
'                    EndLetter strET01, strCP09, "37", strUserNum
'                '2010/11/17 End
                
                'Add By Sindy 2011/2/25
                Case "101" '美國
                    ' 清除定稿例外欄位檔原有資料
                    strET03 = "37" 'Added by Lydia 2024/04/17
                    EndLetter strET01, strCP09, "37", strUserNum
                '2011/2/25 End
                'Add By Sindy 2012/9/5
                Case "239" '歐盟
                    ' 清除定稿例外欄位檔原有資料
                    strET03 = "38" 'Added by Lydia 2024/04/17
                    EndLetter strET01, strCP09, "38", strUserNum
                '2012/9/5 End
                'Add By Sindy 2012/11/9
                Case "011" '日本
                    ' 清除定稿例外欄位檔原有資料
                    strET03 = "39" 'Added by Lydia 2024/04/17
                    EndLetter strET01, strCP09, "39", strUserNum
                '2012/11/9 End
                'Add By Sindy 2013/3/19
                Case "014" '新加坡
                    ' 清除定稿例外欄位檔原有資料
                    strET03 = "40" 'Added by Lydia 2024/04/17
                    EndLetter strET01, strCP09, "40", strUserNum
                '2013/3/19 End
                'add by sonia 2015/11/4
                Case "126" '智利
                    ' 清除定稿例外欄位檔原有資料
                    strET03 = "41" 'Added by Lydia 2024/04/17
                    EndLetter strET01, strCP09, "41", strUserNum
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & strET01 & "','" & strCP09 & "','" & "41" & "','" & strUserNum & _
                             "','延展年度','" & m_NA14 & "')"
                    cnnConnection.Execute strSql
                'end 2015/11/4
                'Add By Sindy 2019/3/22
                Case "104" '墨西哥
                    ' 清除定稿例外欄位檔原有資料
                    strET03 = "42" 'Added by Lydia 2024/04/17
                    EndLetter strET01, strCP09, "42", strUserNum
                '2019/3/22 END
                'Add By Sindy 2021/6/7
                Case "201" '英國
                    ' 清除定稿例外欄位檔原有資料
                    strET03 = "43" 'Added by Lydia 2024/04/17
                    EndLetter strET01, strCP09, "43", strUserNum
                '2021/6/7 END
                'Add By Sindy 2022/1/28
                Case "118" '阿根廷
                    ' 清除定稿例外欄位檔原有資料
                    strET03 = "44" 'Added by Lydia 2024/04/17
                    EndLetter strET01, strCP09, "44", strUserNum
                'add by sonia 2023/9/15
                Case "110" '海地
                    ' 清除定稿例外欄位檔原有資料
                    strET03 = "45" 'Added by Lydia 2024/04/17
                    EndLetter strET01, strCP09, "45", strUserNum
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & strET01 & "','" & strCP09 & "','" & "45" & "','" & strUserNum & _
                             "','附註','" & textTM15S & "')"
                    cnnConnection.Execute strSql
               'end 2023/9/15
                'Added by Lydia 2024/03/19
                Case "044" '澳門
                    ' 清除定稿例外欄位檔原有資料
                    strET03 = "46" 'Added by Lydia 2024/04/17
                    EndLetter strET01, strCP09, "46", strUserNum
                'end 2024/03/19
                Case Else '其他國家
                    If textTM15S = textTM15 Then    '2013/10/11 ADD BY SONIA 判斷延展審定號數與原審定號數是否相同
                        ' 清除定稿例外欄位檔原有資料
                        strET03 = "18" 'Added by Lydia 2024/04/17
                        EndLetter strET01, strCP09, "18", strUserNum
                        'Add By Sindy 2012/11/28
                        If m_TM10 = "113" And Val(DBDATE(textRegDate)) < 20080917 Then '委內瑞拉
                           '註冊日期在2008年9月17日之前者,專用期限自註冊日起算10年
                           '註冊日期在2008年9月17日之後者,專用期限自註冊日起算15年(國家檔設定)
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & strET01 & "','" & strCP09 & "','" & "18" & "','" & strUserNum & _
                                    "','延展年度','10')"
                           cnnConnection.Execute strSql
                        Else
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & strET01 & "','" & strCP09 & "','" & "18" & "','" & strUserNum & _
                                    "','延展年度','" & m_NA14 & "')"
                           cnnConnection.Execute strSql
                        End If
                        '2012/11/28 End
                        'Add By Sindy 2020/7/17
                        If m_TM10 = "034" Then '阿拉伯聯合大公國
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & strET01 & "','" & strCP09 & "','" & "18" & "','" & strUserNum & _
                                    "','顯示申請案號','♀')"
                           cnnConnection.Execute strSql
                        End If
                        '2020/7/17 END
                        'Added by Lydia 2021/03/11 英國寄證書及延展證書定稿增加說明
                        If m_TM10 = "201" Then
                            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                     "VALUES ('" & strET01 & "','" & strCP09 & "','" & "18" & "','" & strUserNum & _
                                     "','特殊字樣','(電子傳送註冊證，英國智慧財產局已停止核發紙本證書)')"
                            cnnConnection.Execute strSql
                        End If
                        'end 2021/03/11
                    '2013/10/11 ADD BY SONIA 判斷延展審定號數與原審定號數不同時
                    Else
                        ' 清除定稿例外欄位檔原有資料
                        strET03 = "17" 'Added by Lydia 2024/04/17
                        EndLetter strET01, strCP09, "17", strUserNum
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & "17" & "','" & strUserNum & _
                                 "','延展年度','" & m_NA14 & "')"
                        cnnConnection.Execute strSql
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & "17" & "','" & strUserNum & _
                                 "','附註','" & textTM15S & "')"
                        cnnConnection.Execute strSql
                    End If
                    '2013/10/11 END
                End Select
                'End
                
                'Added by Lydia 2024/04/17 更改延展證書
                If Text3 = "Y" And strET03 <> "" Then
                  If m_TM10 = "206" Then '奧地利
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                              "','更改延展註冊證','頃接代理人轉來本案之延展核准通知書如檢附，以供查照。因通知書內容有誤，代理人已向當局請求修正，修正尚需一段時間，俟修正完成後，當另行奉達。')"
                     cnnConnection.Execute strSql
                  ElseIf m_TM10 = "016" Then  '紐西蘭
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                              "','更改延展註冊證','惟延展核准證明內容有誤，代理人已將本案延展核准證明交回當局請求修正，修正尚需一段時間，俟收到修正完成之延展核准證明，當另行奉達。隨函暫附延展核准證明影本，以供備查。')"
                     cnnConnection.Execute strSql
                  ElseIf m_TM10 = "015" Then  '澳洲
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                              "','更改延展註冊證','惟延展核准通知書內容有誤，代理人已將本案延展核准通知書交回當局請求修正，修正尚需一段時間，俟收到修正完成之延展核准通知書，當另行奉達。隨函暫附延展核准通知書影本，以供備查。')"
                     cnnConnection.Execute strSql
                  ElseIf InStr("114哥倫比亞,115厄瓜多,116秘魯,120玻利維亞,201英國,126智利,118阿根廷", m_TM10) > 0 Then
                     '電子證書
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                              "','更改延展註冊證','惟延展證書內容有誤，代理人已向當局請求修正，修正尚需一段時間，俟收到正確之延展證書，當另行奉達。隨函暫附電子證書列印本，以供備查。')"
                     cnnConnection.Execute strSql
                  ElseIf m_TM10 = "101" Then  '美國
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                              "','更改延展註冊證','如檢附，以供查照，因通知書內容有誤，代理人已向當局請求修正，修正尚需一段時間，俟修正完成後，當另行奉達。另檢附從美國專利商標局網站列印下來有關本案之進度資料，顯示延展已核准。')"
                     cnnConnection.Execute strSql
                  ElseIf m_TM10 = "239" Then  '歐盟
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                              "','更改延展註冊證','如檢附，以供查照，因通知書內容有誤，代理人已向當局請求修正，修正尚需一段時間，俟修正完成後，當另行奉達。另檢附從歐盟智慧財產局網站列印下來有關本案之進度資料，顯示延展已核准。')"
                     cnnConnection.Execute strSql
                  ElseIf m_TM10 = "044" Then  '澳門
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                              "','更改延展註冊證','惟記錄延展核准之註冊證續頁內容有誤，代理人已交回當局請求修正，修正尚需一段時間，俟收到修正完成之註冊證續頁，當另行奉達。隨函暫附記錄延展核准之註冊證續頁及商標註冊登記簿影本，以供備查。')"
                     cnnConnection.Execute strSql
                  Else   '一般
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                              "','更改延展註冊證','惟延展證書內容有誤，代理人已將本案延展證書交回當局請求修正，修正尚需一段時間，俟收到修正完成之延展證書，當另行奉達。隨函暫附延展證書影本，以供備查。')"
                     cnnConnection.Execute strSql
                  End If
                End If
                'end 2024/04/17
'--------101申請  ---memo by Lydia 2024/04/17
            ' 申請
            'edit by nick 2004/10/21
            'Case "101":
            Case "101", "107", "308":
               Select Case m_TM10
                  ' 烏拉圭
                  Case "121":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "01" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
'                        ' 加註領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(Val(textFee), "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     End If
                  ' 俄羅斯
                  'Modified by Morgan 2017/7/3 俄羅斯5/2改國家代碼 "233"->"023"
                  Case "023":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "02" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
'                        ' 加註領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(Val(textFee), "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     End If
                  ' 哥斯大黎加
                  Case "109":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "03" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
'                        ' 加註領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(Val(textFee), "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     End If
                  ' 尼加拉瓜
                  Case "108":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "04" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
                        ' 加註領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(Val(textFee), "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     End If
                  ' 埃及
                  Case "303":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "05" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
'''                        ' 加註領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(Val(textFee), "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     End If
                  ' 波蘭
                  Case "222":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "06" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
                        ' 加註領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(Val(textFee), "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     End If
                  ' 巴拿馬
                  Case "103":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "07" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
'                        ' 加註領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(Val(textFee), "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     End If
                  ' 柬埔寨
                  Case "046":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "08" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
                        ' 加註領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(Val(textFee), "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
'2011/9/26 CANCEL BY SONIA
'                     Else
'                        ' 法定期限
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','法定期限','" & DBDATE(textTM22) & "')"
'                        cnnConnection.Execute strSql
'2011/9/26 END
                     End If
                  ' 日本
                  Case "011":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "09" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
'                        ' 加註領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(Val(textFee), "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     End If
                  ' 西班牙
                  Case "211":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "10" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
'                        ' 加註領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(Val(textFee), "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     End If
                  ' 葡萄牙
                  Case "213":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "11" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
'                        ' 加註領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(Val(textFee), "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     End If
                  ' 菲律賓
                  Case "030":
                        ' 清除定稿例外欄位檔原有資料
                        strET03 = "12" 'Add By Sindy 2023/5/3
                        EndLetter strET01, strCP09, strET03, strUserNum
                        ' 加註領證費
                        If IsEmptyText(textFee) = False Then
'                           ' 加註領證費
'                           strTemp2 = "本案最後領證程序費用為新台幣" & Format(Val(textFee), "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                    "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                           cnnConnection.Execute strSql
                           'Modify By Sindy 2011/5/24 改為報價通知
                           'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                           'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                           PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                           InsExpField1 strCP09, strLC02, strET03
                           strExc(0) = CompWorkDay(5, strSrvDate(1))
                           strExc(1) = DBDATE(m_NP08)
                           '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                           If Val(strExc(1)) <= Val(strExc(0)) Then
                              PUB_Cache2Letter strCP09, strLC02, False, False
                           End If
                           '2011/5/24 End
'2011/9/26 CANCEL BY SONIA
'                        Else
'                           ' 法定期限
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','法定期限','" & DBDATE(textTM22) & "')"
'                           cnnConnection.Execute strSql
'2011/9/26 END
                        End If
                        'Add By Cheng 2003/04/01
                        ' 商標專用年度
'2009/8/24 modify by sonia 葉易雲說全部改為10年
'                        strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                              "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                              "','商標專用年度','" & Left(DBDATE(textTM22), 4) - Left(DBDATE(textTM21), 4) & "')"
'                        cnnConnection.Execute strSQL
'2009/8/24 end
                  ' 印尼
                  Case "017":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "13" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
'                        ' 加註領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(Val(textFee), "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     End If
                  ' 奈及利亞
                  Case "302":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "14" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     'Add By Sindy 2009/09/30
                     If IsEmptyText(textFee) = False Then
'                        ' 領證費
'                        strTemp2 = "本項最後領證程序費用為新台幣" & Format(textFee.Text, "#,##0") & "元整，敬請撥冗惠付，俾支付代理人費用。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     '2009/09/30 End
                     Else
                        If Val(DBDATE(strSrvDate(1))) > Val(DBDATE(textTM22)) = True Then
                           strTemp1 = "雖本案之專用期限已於" & DBYEAR(textTM22) & "年" & DBMONTH(textTM22) & "月" & DBDAY(textTM22) & "日" & "屆滿，但因係奈國商標局公告作業拖延過久而導致  " & _
                                      "貴公司在專用期限屆滿後才收到註冊證書，故為彌補本身之作業缺失，奈國商標局規定，若註冊專用權人於收到註冊證書後立即辦理專用權延展，則仍可取得十四年之延展" & _
                                      "專用期間; 若未即時辦理延展，商標權即自動失效。因此，本案只要及時辦理延展，仍可享有十四年之延展專用期間。"
                           strTemp2 = "貴公司是否同意續行延展程序，敬請仔細斟酌後，於" & DBYEAR(m_NP08) & "年" & DBMONTH(m_NP08) & "月" & DBDAY(m_NP08) & "日" & "前惠示本所，俾適時請代理人處理。" & _
                                      "至於延展費用為新台幣壹萬陸仟元整，併此說明，以供參酌。"
                           ' 過期商標一
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','過期商標一','" & strTemp1 & "')"
                           cnnConnection.Execute strSql
                           ' 過期商標二
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','過期商標二','" & strTemp2 & vbCrLf & "')"
                           cnnConnection.Execute strSql
                        End If
                     End If
                  ' 厄瓜多, 秘魯, 哥倫比亞, 玻利維亞
                  '2009/3/10 MODIFY BY SONIA 取消委內瑞拉113
                  Case "115", "116", "114", "120":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "15" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     If IsEmptyText(textFee) = False Then
'                        ' 領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(textFee.Text, "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     End If
                     'Add By Sindy 2019/4/10
                     If m_TM10 = "114" Then '哥倫比亞
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                 "','特殊字樣','(電子傳送註冊證，哥倫比亞商標局不核發紙本證書)')"
                        cnnConnection.Execute strSql
                     End If
                  ' 緬甸
                  Case "048":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "16" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     If IsEmptyText(textFee) = False Then
'                        ' 領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(textFee.Text, "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     Else
                        ' 緬甸公報
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                 "','緬甸公報','" & "第" & TextMyanmar_1 & "冊增訂之第" & TextMyanmar_2 & "卷第" & TextMyanmar_3 & "頁" & "')"
                        cnnConnection.Execute strSql
                     End If
                  'Add By Cheng 2002/12/28
                  ' 法國
                  Case "203":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "18" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
'                        ' 加註領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(textFee.Text, "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     Else
                        ' 未使用撤銷說明
                        If IsEmptyText(GetNA19(m_TM10)) = False Then
                           ' 未使用撤銷說明
                           strTemp1 = "依" & GetNationName(m_TM10, 0) & "商標法規定，註冊商標若無正當理由有連續" & GetNA19(m_TM10) & "年以上未使用時，商標有遭撤銷之慮。敬請  貴公司留意本件商標之使用情形，" & _
                                      "使用時務請標示註冊字樣，日後如有侵害情事方得據以要求賠償。此外，本件商標日後如有移轉或註冊事項變更等情事須辦理變更登記。"
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','未使用撤銷說明','" & strTemp1 & "')"
                           cnnConnection.Execute strSql
                        End If
                     End If
                    'Add By Cheng 2003/02/27
                  ' 澳門
                  Case "044":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "20" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
'                        ' 加註領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(textFee.Text, "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     End If
                  'Add By Cheng 2003/03/19
                  '2009/3/18 MODIFY BY SONIA 加319模里西斯(一般但無連續使用限制)
                  'Modify By Sindy 2010/10/15
                  ' 瓜地馬拉
                  'Case "107", "319":
                  Case "107":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "21" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
'                        ' 加註領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(Val(textFee), "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     Else
                        'Add By Cheng 2003/01/01
                        '註冊日期
                        If IsEmptyText(Me.textRegDate.Text) = False Then
                           strTemp2 = "註冊日期為" & Left(DBDATE(Me.textRegDate.Text), 4) & "年" & Mid(DBDATE(Me.textRegDate.Text), 5, 2) & "月" & Right(DBDATE(Me.textRegDate.Text), 2) & "日" & "，"
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','註冊日期','" & strTemp2 & "')"
                           cnnConnection.Execute strSql
                        End If
                     End If
                  'Modify By Sindy 2010/10/15
                  ' 模里西斯
                  Case "319":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "32" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
'                        ' 加註領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(Val(textFee), "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     Else
                        '註冊日期
                        If IsEmptyText(Me.textRegDate.Text) = False Then
                           strTemp2 = "註冊日期為" & Left(DBDATE(Me.textRegDate.Text), 4) & "年" & Mid(DBDATE(Me.textRegDate.Text), 5, 2) & "月" & Right(DBDATE(Me.textRegDate.Text), 2) & "日" & "，"
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','註冊日期','" & strTemp2 & "')"
                           cnnConnection.Execute strSql
                        End If
                     End If
                  '92.10.14 ADD BY SONIA
                  ' 敘利亞
                  Case "032":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "23" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
'                        ' 加註領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(Val(textFee), "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     Else
                        'Add By Cheng 2003/01/01
                        '註冊日期
                        If IsEmptyText(Me.textRegDate.Text) = False Then
                           strTemp2 = "註冊日期為" & Left(DBDATE(Me.textRegDate.Text), 4) & "年" & Mid(DBDATE(Me.textRegDate.Text), 5, 2) & "月" & Right(DBDATE(Me.textRegDate.Text), 2) & "日" & "，"
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','註冊日期','" & strTemp2 & "')"
                           cnnConnection.Execute strSql
                        End If
                     End If
                'Add By Cheng 2004/04/27
                  Case "126": '智利
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "25" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
'                        ' 加註領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(Val(textFee), "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     Else
                        ' 未使用撤銷說明
                        If IsEmptyText(GetNA19(m_TM10)) = False Then
                           ' 未使用撤銷說明
                           strTemp1 = "依" & GetNationName(m_TM10, 0) & "商標法規定，註冊商標若無正當理由有連續" & GetNA19(m_TM10) & "年以上未使用時，商標有遭撤銷之慮。敬請  貴公司留意本件商標之使用情形，" & _
                                      "使用時務請標示註冊字樣，日後如有侵害情事方得據以要求賠償。此外，本件商標日後如有移轉或註冊事項變更等情事須辦理變更登記。"
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','未使用撤銷說明','" & strTemp1 & "')"
                           cnnConnection.Execute strSql
                        End If
                        'Add By Cheng 2003/01/01
                        '註冊日期
                        If IsEmptyText(Me.textRegDate.Text) = False Then
                           strTemp2 = "註冊日期為" & Left(DBDATE(Me.textRegDate.Text), 4) & "年" & Mid(DBDATE(Me.textRegDate.Text), 5, 2) & "月" & Right(DBDATE(Me.textRegDate.Text), 2) & "日" & "，"
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','註冊日期','" & strTemp2 & "')"
                           cnnConnection.Execute strSql
                        End If
                     End If
                  'add by nickc 2006/02/21 葉芳如 改美國註冊証定稿
                  Case "101":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "27" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
'                        ' 加註領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(Val(textFee), "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     Else
                        ' 未使用撤銷說明
                        If IsEmptyText(GetNA19(m_TM10)) = False Then
                           ' 未使用撤銷說明
                           strTemp1 = "依" & GetNationName(m_TM10, 0) & "商標法規定，註冊商標若無正當理由有連續" & GetNA19(m_TM10) & "年以上未使用時，商標有遭撤銷之慮。敬請  貴公司留意本件商標之使用情形，" & _
                                      "使用時務請標示註冊字樣，日後如有侵害情事方得據以要求賠償。此外，本件商標日後如有移轉或註冊事項變更等情事須辦理變更登記。"
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','未使用撤銷說明','" & strTemp1 & "')"
                           cnnConnection.Execute strSql
                        End If
                        'Add By Cheng 2003/01/01
                        '註冊日期
                        If IsEmptyText(Me.textRegDate.Text) = False Then
                           strTemp2 = "註冊日期為" & Left(DBDATE(Me.textRegDate.Text), 4) & "年" & Mid(DBDATE(Me.textRegDate.Text), 5, 2) & "月" & Right(DBDATE(Me.textRegDate.Text), 2) & "日" & "，"
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','註冊日期','" & strTemp2 & "')"
                           cnnConnection.Execute strSql
                        End If
                        'Add By nickc 2006/02/21
                        '提出宣誓日
                        If IsEmptyText(Me.textRegDate.Text) = False Then
                           strTemp2 = " " & Val(Left(DBDATE(Me.textRegDate.Text), 4)) + 5 & " 年 " & Val(Mid(DBDATE(Me.textRegDate.Text), 5, 2)) & " 月 " & Val(Right(DBDATE(Me.textRegDate.Text), 2)) & " 日"
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','提出宣誓日','" & strTemp2 & "')"
                           cnnConnection.Execute strSql
                        End If
                     End If
                  'add by nickc 2007/04/25 加入非洲聯盟
                  Case "304"
                     strET03 = "22" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
'                        ' 加註領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(Val(textFee), "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     End If
                  '2009/6/1 add by sonia 葉易雲 改尼泊爾註冊証定稿,專用期間屆滿前35日內辦理延展
                  Case "058"
                     strET03 = "28" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
'                        ' 加註領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(Val(textFee), "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     Else
                        'Add By Cheng 2003/01/01
                        '註冊日期
                        If IsEmptyText(Me.textRegDate.Text) = False Then
                           strTemp2 = "註冊日期為" & Left(DBDATE(Me.textRegDate.Text), 4) & "年" & Mid(DBDATE(Me.textRegDate.Text), 5, 2) & "月" & Right(DBDATE(Me.textRegDate.Text), 2) & "日" & "，"
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','註冊日期','" & strTemp2 & "')"
                           cnnConnection.Execute strSql
                        End If
                     End If
                  '2009/6/1 end
                  'Add By Sindy 2010/6/3
                  ' 歐盟
                  Case "239":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "29" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
'                        ' 加註領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(Val(textFee), "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     Else
                        '註冊日期
                        If IsEmptyText(Me.textRegDate.Text) = False Then
                           strTemp2 = "註冊日期為" & Left(DBDATE(Me.textRegDate.Text), 4) & "年" & Mid(DBDATE(Me.textRegDate.Text), 5, 2) & "月" & Right(DBDATE(Me.textRegDate.Text), 2) & "日" & "，"
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','註冊日期','" & strTemp2 & "')"
                           cnnConnection.Execute strSql
                        End If
                     End If
                  'Add By Sindy 2010/7/6
                  ' 挪威
                  Case "215":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "30" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
'                        ' 加註領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(Val(textFee), "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     Else
                        '註冊日期
                        If IsEmptyText(Me.textRegDate.Text) = False Then
                           strTemp2 = "註冊日期為" & Left(DBDATE(Me.textRegDate.Text), 4) & "年" & Mid(DBDATE(Me.textRegDate.Text), 5, 2) & "月" & Right(DBDATE(Me.textRegDate.Text), 2) & "日" & "，"
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','註冊日期','" & strTemp2 & "')"
                           cnnConnection.Execute strSql
                        End If
                     End If
                  'Add By Sindy 2010/7/15
                  ' 瑞士
                  Case "205":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "31" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
'                        ' 加註領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(Val(textFee), "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     Else
                        '註冊日期
                        If IsEmptyText(Me.textRegDate.Text) = False Then
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','註冊日期','" & DBDATE(Me.textRegDate.Text) & "')"
                           cnnConnection.Execute strSql
                        End If
                        '註冊日期(迄)
                        If IsEmptyText(Me.textRegDate.Text) = False Then
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','註冊日期(迄)','" & DBDATE(DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(Me.textRegDate.Text)))) & "')"
                           cnnConnection.Execute strSql
                        End If
                     End If
                  'Add By Sindy 2011/6/29
                  ' 德國
                  Case "231":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "33" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
                        '報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                     Else
                        '註冊日期
                        If IsEmptyText(Me.textRegDate.Text) = False Then
                           strTemp2 = "註冊日期為" & Left(DBDATE(Me.textRegDate.Text), 4) & "年" & Mid(DBDATE(Me.textRegDate.Text), 5, 2) & "月" & Right(DBDATE(Me.textRegDate.Text), 2) & "日" & "，"
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','註冊日期','" & strTemp2 & "')"
                           cnnConnection.Execute strSql
                        End If
                     End If
                  'Add By Sindy 2012/6/22
                  ' 加拿大
                  Case "102":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "34" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                     Else
                        '註冊日期
                        If IsEmptyText(Me.textRegDate.Text) = False Then
                           strTemp2 = "註冊日期為" & Left(DBDATE(Me.textRegDate.Text), 4) & "年" & Mid(DBDATE(Me.textRegDate.Text), 5, 2) & "月" & Right(DBDATE(Me.textRegDate.Text), 2) & "日" & "，"
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','註冊日期','" & strTemp2 & "')"
                           cnnConnection.Execute strSql
                        End If
                     End If
                     'add by sonia 2019/6/19
                     If Val(DBDATE(textRegDate)) <= 20190617 Then '修法
                        '發證日在2019/06/17以前者維持15年,之後改為10年(國家檔設定)；每次延展一律改為10年(國家檔設定)(原為15年)
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                 "','商標專用年度','15')"
                        cnnConnection.Execute strSql
                     Else
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                 "','商標專用年度','" & NickTmNa13 & "')"
                        cnnConnection.Execute strSql
                     End If
                     'end 2019/6/19
                  'add by sonia 2014/10/31
                  ' 莫三比克
                  Case "318":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "35" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                              "','延展年度','" & m_NA14 & "')"
                     cnnConnection.Execute strSql
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                              "','商標專用年度','" & NickTmNa13 & "')"
                     cnnConnection.Execute strSql
                   ' 加註領證費
                     If IsEmptyText(textFee) = False Then
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                     Else
                        '註冊日期
                        If IsEmptyText(Me.textRegDate.Text) = False Then
                           strTemp2 = "註冊日期為" & Left(DBDATE(Me.textRegDate.Text), 4) & "年" & Mid(DBDATE(Me.textRegDate.Text), 5, 2) & "月" & Right(DBDATE(Me.textRegDate.Text), 2) & "日" & "，"
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','註冊日期','" & strTemp2 & "')"
                           cnnConnection.Execute strSql
                        End If
                     End If
                  'end 2014/10/31
                  'Add By Sindy 2015/1/27
                  Case "014" '新加坡
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "36" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                              "','延展年度','" & m_NA14 & "')"
                     cnnConnection.Execute strSql
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                              "','商標專用年度','" & NickTmNa13 & "')"
                     cnnConnection.Execute strSql
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     Else
                        '註冊日期
                        If IsEmptyText(Me.textRegDate.Text) = False Then
                           strTemp2 = "註冊日期為" & Left(DBDATE(Me.textRegDate.Text), 4) & "年" & Mid(DBDATE(Me.textRegDate.Text), 5, 2) & "月" & Right(DBDATE(Me.textRegDate.Text), 2) & "日" & "，"
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','註冊日期','" & strTemp2 & "')"
                           cnnConnection.Execute strSql
                        End If
                     End If
                  '2015/1/27 END
                  '2016/8/19 add by sonia
                  Case "112" '波多黎各
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "37" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                     End If
                  'end 2016/8/19
                  ' 墨西哥  add by sonia 2018/11/19
                  Case "104":
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "39" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     strTemp2 = "註冊日期為" & Left(DBDATE(Me.textRegDate.Text), 4) & "年" & Mid(DBDATE(Me.textRegDate.Text), 5, 2) & "月" & Right(DBDATE(Me.textRegDate.Text), 2) & "日" & "，"
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                              "','註冊日期','" & strTemp2 & "')"
                     cnnConnection.Execute strSql
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                              "','延展年度','" & m_NA14 & "')"
                     cnnConnection.Execute strSql
                   ' 加註領證費
                     If IsEmptyText(textFee) = False Then
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                     End If
                  'end 2018/11/19
                  'add by sonia  2022/10/14 通用定稿改專用定稿
                  Case "118" '阿根廷
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "44"
                     EndLetter strET01, strCP09, strET03, strUserNum
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                     End If
                  ' 其它
                  Case Else:
                     ' 清除定稿例外欄位檔原有資料
                     strET03 = "17" 'Add By Sindy 2023/5/3
                     EndLetter strET01, strCP09, strET03, strUserNum
                     'Add By Sindy 2012/11/28
                     If m_TM10 = "113" And Val(DBDATE(textRegDate)) < 20080917 Then '委內瑞拉
                        '註冊日期在2008年9月17日之前者,專用期限自註冊日起算10年
                        '註冊日期在2008年9月17日之後者,專用期限自註冊日起算15年(國家檔設定)
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                 "','延展年度','10')"
                        cnnConnection.Execute strSql
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                 "','商標專用年度','10')"
                        cnnConnection.Execute strSql
                     Else
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                 "','延展年度','" & m_NA14 & "')"
                        cnnConnection.Execute strSql
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                 "','商標專用年度','" & NickTmNa13 & "')"
                        cnnConnection.Execute strSql
                     End If
                     '2012/11/28 End
                     ' 加註領證費
                     If IsEmptyText(textFee) = False Then
'                        ' 加註領證費
'                        strTemp2 = "本案最後領證程序費用為新台幣" & Format(Val(textFee), "#,##0") & "元整，敬請撥冗惠付！不勝感荷。"
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','加註領證費','" & strTemp2 & vbCrLf & "')"
'                        cnnConnection.Execute strSql
                        'Modify By Sindy 2011/5/24 改為報價通知
                        'Modify By Sindy 2012/6/12 把 strCP09 改為 m_NP01
                        'Modify By Sindy 2015/3/10 把 m_NP01 改為 strCP09
                        PUB_AddLetterCache strCP09, strLC02, strCP09, strET01, strET03
                        InsExpField1 strCP09, strLC02, strET03
                        strExc(0) = CompWorkDay(5, strSrvDate(1))
                        strExc(1) = DBDATE(m_NP08)
                        '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                        If Val(strExc(1)) <= Val(strExc(0)) Then
                           PUB_Cache2Letter strCP09, strLC02, False, False
                        End If
                        '2011/5/24 End
                     Else
                        'Add By Cheng 2003/01/01
                        '註冊日期
                        If IsEmptyText(Me.textRegDate.Text) = False Then
                           strTemp2 = "註冊日期為" & Left(DBDATE(Me.textRegDate.Text), 4) & "年" & Mid(DBDATE(Me.textRegDate.Text), 5, 2) & "月" & Right(DBDATE(Me.textRegDate.Text), 2) & "日" & "，"
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                    "','註冊日期','" & strTemp2 & "')"
                           cnnConnection.Execute strSql
                        End If
                     End If
                     'Add By Sindy 2016/6/29
                     If m_TM10 = "235" Then '土耳其
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                 "','特殊字樣','(電子傳送註冊證，土耳其專利局不核發紙本證書)及其英譯文')"
                        cnnConnection.Execute strSql
'cancel by sonia 2023/12/12 改用專用定稿
'                     'Add By Sindy 2017/2/24
'                     ElseIf m_TM10 = "118" Then '阿根廷
'                        'Modify By Sindy 2018/12/20
'                        '電子傳送註冊證，阿根廷工業財產局不核發紙本證書 => 電子傳送註冊證，阿根廷商標局不核發紙本證書
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
'                                 "','特殊字樣','(電子傳送註冊證，阿根廷商標局不核發紙本證書)')"
'                        cnnConnection.Execute strSql
'end 2023/12/12
                     'Add By Sindy 2019/4/10
                     ElseIf m_TM10 = "117" Then '巴西
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                 "','特殊字樣','(電子傳送註冊證，巴西工業財產局不核發紙本證書)')"
                        cnnConnection.Execute strSql
                     'Add By Sindy 2020/1/10 015.澳洲
                     'Add By Sindy 2020/7/3 016.紐西蘭
                     'Add By Sindy 2020/10/23 040.印度
                     'Modified by Lydia 2024/03/19 +301.南非
                     ElseIf m_TM10 = "015" Or m_TM10 = "016" Or m_TM10 = "040" Or m_TM10 = "301" Then
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                 "','電子','電子')"
                        cnnConnection.Execute strSql
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                 "','特殊字樣','(" & textTM10 & "不核發紙本證書)')"
                        cnnConnection.Execute strSql
                     'Added by Lydia 2021/03/11 英國寄證書及延展證書定稿增加說明
                     ElseIf m_TM10 = "201" Then
                            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                     "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                     "','特殊字樣','(電子傳送註冊證，英國智慧財產局已停止核發紙本證書)')"
                            cnnConnection.Execute strSql
                     'end 2021/03/11
                     'add by sonia 2024/1/3 越南
                     ElseIf m_TM10 = "042" Then
                            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                     "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                     "','特殊字樣','及其英譯文')"
                            cnnConnection.Execute strSql
                     'end 2024/1/3
                     End If
                     'Add By Sindy 2020/8/24
                     If m_TM10 = "326" Then '辛巴威
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                 "','各國專屬提醒','若有授權他人在辛巴威使用本商標，亦需向官方辦理登記，才能有效對抗第三人。')"
                        cnnConnection.Execute strSql
                     ElseIf m_TM10 = "321" Then '馬達加斯加
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                 "','各國專屬提醒','若有授權他人在馬達加斯加使用本商標，亦需向官方辦理登記，才能有效對抗第三人。')"
                        cnnConnection.Execute strSql
                     End If
                     '2020/8/24 END
                     'Add By Sindy 2023/2/13 各國加註
                     If m_TM10 = "028" Then '科威特
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                 "','各國加註','依科威特現行商標實務，商品/服務僅需指定類別，不需指定商品/服務內容，因此，本件商標之商品/服務範圍係包括本類所有商品/服務，併此說明。" & Chr(13) & Chr(10) & "')"
                        cnnConnection.Execute strSql
                     End If
                     '2020/8/24 END
               End Select
               'Add By Sindy 2023/5/3
               If Text3 = "Y" And strET03 <> "" Then
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & strET01 & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                           "','更改註冊證','惟註冊證內容有誤，代理人已將本案註冊證交回當局請求修正，修正尚需一段時間，俟收到修正完成之註冊證正本，當另行奉達。隨函暫附註冊證影本，以供備查。')"
                  cnnConnection.Execute strSql
               End If
         End Select
   End Select
End Sub

'Add By Sindy 2011/5/24
'寫例外欄位到暫存檔
Private Sub InsExpField1(NP01 As String, NP22 As String, Optional ET03 As String)
Dim strTemp1 As String, strTemp2 As String
   
   strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                   "VALUES ('" & NP01 & "'," & NP22 & ",'領證費','" & Me.textFee.Text & "','Y')"
   cnnConnection.Execute strSql
   
   strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                   "VALUES ('" & NP01 & "'," & NP22 & ",'領證費點數','" & Me.textCP18.Text & "','')"
   cnnConnection.Execute strSql
   
   'add by sonia 2014/6/27 CFT-015174
   strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                   "VALUES ('" & NP01 & "'," & NP22 & ",'延展年度','" & m_NA14 & "','')"
   cnnConnection.Execute strSql
   
   'Added by Lydia 2021/04/07 英國寄證書及延展證書定稿增加說明; CFT-21948採用報價轉定稿,所以也要增加定稿暫存檔
   If m_TM10 = "201" And (ET03 = "17" Or ET03 = "18") Then
        strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                        "VALUES ('" & NP01 & "'," & NP22 & ",'特殊字樣','(電子傳送註冊證，英國智慧財產局已停止核發紙本證書)','')"
        cnnConnection.Execute strSql
   End If
   'end 2021/04/07
   
   If ChkTG = True Then
      strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                      "VALUES ('" & NP01 & "'," & NP22 & ",'ChkTG','True-0526','')"
      cnnConnection.Execute strSql
   End If
      
'2011/9/26 CANCEL BY SONIA
'   If ET03 = "08" Or ET03 = "12" Then
'      strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
'                      "VALUES ('" & NP01 & "'," & NP22 & ",'法定期限','" & DBDATE(textTM22) & "','')"
'      cnnConnection.Execute strSql
'   End If
'2011/9/26 END

   If ET03 = "14" Then
      If Val(DBDATE(strSrvDate(1))) > Val(DBDATE(textTM22)) = True Then
         strTemp1 = "雖本案之專用期限已於" & DBYEAR(textTM22) & "年" & DBMONTH(textTM22) & "月" & DBDAY(textTM22) & "日" & "屆滿，但因係奈國商標局公告作業拖延過久而導致  " & _
                    "貴公司在專用期限屆滿後才收到註冊證書，故為彌補本身之作業缺失，奈國商標局規定，若註冊專用權人於收到註冊證書後立即辦理專用權延展，則仍可取得十四年之延展" & _
                    "專用期間; 若未即時辦理延展，商標權即自動失效。因此，本案只要及時辦理延展，仍可享有十四年之延展專用期間。"
         strTemp2 = "    貴公司是否同意續行延展程序，敬請仔細斟酌後，於" & DBYEAR(m_NP08) & "年" & DBMONTH(m_NP08) & "月" & DBDAY(m_NP08) & "日" & "前惠示本所，俾適時請代理人處理。" & _
                    "至於延展費用為新台幣壹萬陸仟元整，併此說明，以供參酌。"
         ' 過期商標一
         strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                         "VALUES ('" & NP01 & "'," & NP22 & ",'過期商標一','" & strTemp1 & "','')"
         cnnConnection.Execute strSql
         ' 過期商標二
         strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                         "VALUES ('" & NP01 & "'," & NP22 & ",'過期商標二','" & strTemp2 & vbCrLf & "','')"
         cnnConnection.Execute strSql
      End If
   End If
   If ET03 = "16" Then
      ' 緬甸公報
      strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                      "VALUES ('" & NP01 & "'," & NP22 & ",'緬甸公報','" & "第" & TextMyanmar_1 & "冊增訂之第" & TextMyanmar_2 & "卷第" & TextMyanmar_3 & "頁" & "','')"
      cnnConnection.Execute strSql
   End If
   If ET03 = "18" Or ET03 = "25" Or ET03 = "27" Then
      If IsEmptyText(GetNA19(m_TM10)) = False Then
         ' 未使用撤銷說明
         strTemp1 = "依" & GetNationName(m_TM10, 0) & "商標法規定，註冊商標若無正當理由有連續" & GetNA19(m_TM10) & "年以上未使用時，商標有遭撤銷之慮。敬請  貴公司留意本件商標之使用情形，" & _
                    "使用時務請標示註冊字樣，日後如有侵害情事方得據以要求賠償。此外，本件商標日後如有移轉或註冊事項變更等情事須辦理變更登記。"
         strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                        "VALUES ('" & NP01 & "'," & NP22 & ",'未使用撤銷說明','" & strTemp1 & "','')"
         cnnConnection.Execute strSql
      End If
   End If
   'If ET03 = "21" Or ET03 = "23" Or ET03 = "25" Or ET03 = "27" Or ET03 = "28" Or ET03 = "29" Or ET03 = "30" Or ET03 = "32" Then
   If ET03 = "31" Then
      '註冊日期
      If IsEmptyText(Me.textRegDate.Text) = False Then
         strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                        "VALUES ('" & NP01 & "'," & NP22 & ",'註冊日期','" & DBDATE(Me.textRegDate.Text) & "','')"
         cnnConnection.Execute strSql
      End If
   Else
      If IsEmptyText(Me.textRegDate.Text) = False Then
         '註冊日期
         strTemp2 = "註冊日期為" & Left(DBDATE(Me.textRegDate.Text), 4) & "年" & Mid(DBDATE(Me.textRegDate.Text), 5, 2) & "月" & Right(DBDATE(Me.textRegDate.Text), 2) & "日" & "，"
         strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                        "VALUES ('" & NP01 & "'," & NP22 & ",'註冊日期','" & strTemp2 & "','')"
         cnnConnection.Execute strSql
      End If
   End If
   If ET03 = "27" Then
      If IsEmptyText(Me.textRegDate.Text) = False Then
         '提出宣誓日
         strTemp2 = " " & Val(Left(DBDATE(Me.textRegDate.Text), 4)) + 5 & " 年 " & Val(Mid(DBDATE(Me.textRegDate.Text), 5, 2)) & " 月 " & Val(Right(DBDATE(Me.textRegDate.Text), 2)) & " 日"
         strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                        "VALUES ('" & NP01 & "'," & NP22 & ",'提出宣誓日','" & strTemp2 & "','')"
         cnnConnection.Execute strSql
      End If
   End If
   If ET03 = "31" Then
      '註冊日期(迄)
      If IsEmptyText(Me.textRegDate.Text) = False Then
         strSql = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                        "VALUES ('" & NP01 & "'," & NP22 & ",'註冊日期(迄)','" & DBDATE(DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(Me.textRegDate.Text)))) & "','')"
         cnnConnection.Execute strSql
      End If
   End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
   strET01 = "05" 'Add By Sindy 2023/5/3
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   Select Case m_TM01
      Case "CFC":
         If IsCustomerIndividual(m_TM23) = True Then
            NowPrint strCP09, strET01, "21", False, strUserNum, 0
         Else
            If m_SP48 = "Y" Then
               NowPrint strCP09, strET01, "19", False, strUserNum, 0
            Else
               NowPrint strCP09, strET01, "20", False, strUserNum, 0
            End If
         End If
      Case "CFT":
         
         Select Case m_CP10
            ' 延展
            Case "102":
                'Modify By Cheng 2004/03/10
                Select Case m_TM10
                Case "121" '烏拉圭
                    NowPrint strCP09, strET01, "19", False, strUserNum, 0
                Case "206" '奧地利
                    NowPrint strCP09, strET01, "20", False, strUserNum, 0
                'add by nick 2004/07/05 葉門
                Case "036"
                    NowPrint strCP09, strET01, "21", False, strUserNum, 0
                '2005/5/23 ADD BY SONIA 印度
                Case "040"
                    '2006/3/8 MODIFY BY SONIA MAY說取消此控制,延展全部十年
                    'If m_TM11 < 20010511 Then
                    '   NowPrint strCP09, strET01, "27", False, strUserNum, 0
                    'Else
                    '   NowPrint strCP09, strET01, "18", False, strUserNum, 0
                    'End If
                    NowPrint strCP09, strET01, "18", False, strUserNum, 0
                    '2006/3/8 END
                '2005/5/23 END
                '2006/4/11 ADD BY SONIA 緬甸
                Case "048"
                    'Modified by Lydia 2024/04/17 更正延展證書>> 改為一般
                    'NowPrint strCP09, strET01, "28", False, strUserNum, 0
                    If textTM15S = textTM15 Then
                       NowPrint strCP09, strET01, "18", False, strUserNum, 0
                    Else
                       NowPrint strCP09, strET01, "17", False, strUserNum, 0
                    End If
                    'end 2024/04/17
                '2006/4/11 END
                'add by nickc 2006/11/13 柬埔寨
                Case "046"
                    NowPrint strCP09, strET01, "29", False, strUserNum, 0
                'add by nickc 2007/03/03 加入紐西蘭
                Case "016"
                    NowPrint strCP09, strET01, "30", False, strUserNum, 0
                'add by nickc 2007/06/15  加入澳洲
                Case "015"
                    NowPrint strCP09, strET01, "31", False, strUserNum, 0
                'add by nickc 2007/07/10 加入韓國
                Case "012"
                    NowPrint strCP09, strET01, "32", False, strUserNum, 0
                'add by nickc 2007/08/28 加入波蘭
                Case "222"
                    NowPrint strCP09, strET01, "33", False, strUserNum, 0
                'add by nickc 2008/05/02
                '2009/3/10 MODIFY BY SONIA 取消委內瑞拉113
                Case "120", "115", "114", "116" '玻利維亞、厄瓜多、哥倫比亞、祕魯
                    '2013/10/11 modify by sonia 區分延展號數與原審定號數是否相同
                    'NowPrint strCP09, strET01, "34", False, strUserNum, 0
                    If textTM15S = textTM15 Then
                       NowPrint strCP09, strET01, "34", False, strUserNum, 0
                    Else
                       NowPrint strCP09, strET01, "24", False, strUserNum, 0
                    End If
                    '2013/10/11 end
                
                'Add By Sindy 2009/04/29
                Case "030" '菲賓律
                    NowPrint strCP09, strET01, "35", False, strUserNum, 0
                '2009/04/29 End
                
                'Add By Sindy 2010/4/1
                Case "019" '泰國
                    NowPrint strCP09, strET01, "36", False, strUserNum, 0
                '2010/4/1 End
                
                'Modify By Sindy 2010/12/31 印尼與其他國家只差證書2字, 加拿大也要將核准通知改為證書2字, 因此直接改其他國家定稿, 印尼也不須另立定稿
'                'Add By Sindy 2010/11/17
'                Case "017": '印尼
'                    NowPrint strCP09, strET01, "37", False, strUserNum, 0
'                '2010/11/17 End
                
                'Add By Sindy 2011/2/25
                Case "101" '美國
                    NowPrint strCP09, strET01, "37", False, strUserNum, 0
                '2011/2/25 End
                'Add By Sindy 2012/9/5
                Case "239" '歐盟
                    NowPrint strCP09, strET01, "38", False, strUserNum, 0
                '2012/9/5 End
                'Add By Sindy 2012/11/9
                Case "011" '日本
                    NowPrint strCP09, strET01, "39", False, strUserNum, 0
                '2012/11/9 End
                'Add By Sindy 2013/3/19
                Case "014" '新加坡
                    NowPrint strCP09, strET01, "40", False, strUserNum, 0
                '2013/3/19 End
                'add by sonia 2015/11/4
                Case "126" '智利
                    NowPrint strCP09, strET01, "41", False, strUserNum, 0
                'end 2015/11/4
                'Add By Sindy 2019/3/22
                Case "104" '墨西哥
                    NowPrint strCP09, strET01, "42", False, strUserNum, 0
                '2019/3/22 END
                'Add By Sindy 2021/6/7
                Case "201" '英國
                    NowPrint strCP09, strET01, "43", False, strUserNum, 0
                'Add By Sindy 2022/1/28
                Case "118" '阿根廷
                    NowPrint strCP09, strET01, "44", False, strUserNum, 0
                'add by sonia 2023/9/15
                Case "110" '海地
                    NowPrint strCP09, strET01, "45", False, strUserNum, 0
                'Added by Lydia 2024/03/19
                Case "044" '澳門
                    NowPrint strCP09, strET01, "46", False, strUserNum, 0
                'end 2024/03/19
                Case Else '其他國家
                    '2013/10/11 modify by sonia 區分延展號數與原審定號數是否相同
                    'NowPrint strCP09, strET01, "18", False, strUserNum, 0
                    If textTM15S = textTM15 Then
                       NowPrint strCP09, strET01, "18", False, strUserNum, 0
                    Else
                       NowPrint strCP09, strET01, "17", False, strUserNum, 0
                    End If
                    '2013/10/11 end
                End Select
                'End
            ' 申請
            'edit by nick 2004/10/21
            'Case "101":
            Case "101", "107", "308":
               Select Case m_TM10
                  ' 烏拉圭
                  Case "121":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "01", False, strUserNum, 0
                  ' 俄羅斯
                  'Modified by Morgan 2017/7/3 俄羅斯5/2改國家代碼 "233"->"023"
                  Case "023":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "02", False, strUserNum, 0
                  ' 哥斯大黎加
                  Case "109":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "03", False, strUserNum, 0
                  ' 尼加拉瓜
                  Case "108":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "04", False, strUserNum, 0
                  ' 埃及
                  Case "303":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "05", False, strUserNum, 0
                  ' 波蘭
                  Case "222":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "06", False, strUserNum, 0
                  ' 巴拿馬
                  Case "103":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "07", False, strUserNum, 0
                  ' 柬埔寨
                  Case "046":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "08", False, strUserNum, 0
                  ' 日本
                  Case "011":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "09", False, strUserNum, 0
                  ' 西班牙
                  Case "211":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "10", False, strUserNum, 0
                  ' 葡萄牙
                  Case "213":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "11", False, strUserNum, 0
                  ' 菲律賓
                  Case "030":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "12", False, strUserNum, 0
                  ' 印尼
                  Case "017":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "13", False, strUserNum, 0
                  ' 奈及利亞
                  Case "302":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "14", False, strUserNum, 0
                  ' 厄瓜多, 秘魯, 哥倫比亞, 委內瑞拉, 玻利維亞
                  '2009/3/3 MODIFY BY SONIA 取消委內瑞拉113
                  Case "115", "116", "114", "120":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "15", False, strUserNum, 0
                  ' 緬甸
                  Case "048":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "16", False, strUserNum, 0
                  'Add By Cheng 2002/12/28
                  ' 法國
                  Case "203":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "18", False, strUserNum, 0
                  'Add By Cheng 2003/02/27
                  ' 澳門
                  Case "044":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "20", False, strUserNum, 0
                    'Add By Cheng 2003/02/27
                  '2009/3/18 MODIFY BY SONIA 加319模里西斯(一般但無連續使用限制)
                  'Modify By Sindy 2010/10/15
                  'Case "107", "319":
                  ' 瓜地馬拉
                  Case "107":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "21", False, strUserNum, 0
                  'Modify By Sindy 2010/10/15
                  ' 模里西斯
                  Case "319":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "32", False, strUserNum, 0
                  '92.5.20 Add By SONIA
                  ' 非洲聯盟
                  Case "304":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "22", False, strUserNum, 0
                  '92.10.14 Add By SONIA
                  ' 敘利亞
                  Case "032":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "23", False, strUserNum, 0
                  '92.11.9 Add By SONIA
                  '2008/11/24 MODIFY BY SONIA 加入巴基斯坦038
                  '2009/3/27 MODIFY BY SONIA 葉易雲說2004年4月12日前期滿應延展案件為15年,之後為10年,故取消巴基斯坦038
                  ' 馬來西亞
                  Case "018"
                     If textTM22 <> New_TM22 Then
                        NowPrint strCP09, strET01, "24", False, strUserNum, 0
                     Else
                        NowPrint strCP09, strET01, "17", False, strUserNum, 0
                     End If
                    'Add By Cheng 2004/04/27
                  Case "126": '智利
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "25", False, strUserNum, 0
                  'add by nickc 2006/02/21 葉芳如 改美國註冊証定稿
                  Case "101":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "27", False, strUserNum, 0
                  '2009/6/1 add by sonia 葉易雲 改尼泊爾註冊証定稿,專用期間屆滿前35日內辦理延展
                  Case "058":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "28", False, strUserNum, 0
                  'Add By Sindy 2010/6/3
                  ' 歐盟
                  Case "239":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "29", False, strUserNum, 0
                  'Add By Sindy 2010/7/6
                  ' 挪威
                  Case "215":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "30", False, strUserNum, 0
                  'Add By Sindy 2010/7/15
                  ' 瑞士
                  Case "205":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "31", False, strUserNum, 0
                  'Add By Sindy 2011/6/29
                  ' 德國
                  Case "231":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "33", False, strUserNum, 0
                  'Add By Sindy 2012/6/22
                  ' 加拿大
                  Case "102":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "34", False, strUserNum, 0
                  'add by sonia 2014/10/31
                  ' 莫三比克
                  Case "318":
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "35", False, strUserNum, 0
                  'end 2014/10/31
                  'Add By Sindy 2015/1/27
                  Case "014" '新加坡
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "36", False, strUserNum, 0
                  '2015/1/27 END
                  'add by sonia 2016/1/19
                  Case "112" '波多黎各
                      If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "37", False, strUserNum, 0
                  'end 2016/1/19
                  'add by sonia 2018/11/16
                  Case "104" '墨西哥
                     If IsEmptyText(textFee) = True Then   'add by sonia 2022/1/21 改為報價定稿
                        If m_CP10 = "101" Then
                          NowPrint strCP09, strET01, "39", False, strUserNum, 0
                        Else
                          NowPrint strCP09, strET01, "17", False, strUserNum, 0
                        End If
                     End If                                'add by sonia 2022/1/21 改為報價定稿
                  'end 2018/11/16
                  'add by sonia 2022/10/14 通用定稿改專用定稿
                  Case "118" '阿根廷
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "44", False, strUserNum, 0
                  ' 其它
                  Case Else:
                     If IsEmptyText(textFee) = True Then NowPrint strCP09, strET01, "17", False, strUserNum, 0
               End Select
         End Select
         'add by nick 2004/10/05  商品及服務定稿
         'Modify By Sindy 2011/5/24 若有輸入領證費時則採報價通知流程
         'MODIFY BY SONIA 2014/7/30 非延展案才要,故加入m_CP10控制CFT-006188延展證書不印
         'modify by sonia 2025/9/4 +109緩審延展CFT-016520
         If ChkTG = True And IsEmptyText(textFee) = True And m_CP10 <> "102" And m_CP10 <> "109" Then
            'Add By Sindy 2017/4/18
            Select Case m_TM10
               ' 緬甸
               Case "048":
                  NowPrint strCP09, strET01, "38", False, strUserNum, 0
            '2017/4/18 END
               Case Else
                  NowPrint strCP09, strET01, "26", False, strUserNum, 0
            End Select
         End If
   End Select
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   
   'Add By Sindy 2010/12/24
   If Me.textTM15.Enabled = True Then
      Cancel = False
      textTM15_Validate Cancel
      If Cancel = True Then
         textTM15.SetFocus
         Exit Function
      End If
   End If
   
   If Me.textCP18.Enabled = True Then
      Cancel = False
      textCP18_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textFeeDate.Enabled = True Then
      Cancel = False
      textFeeDate_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textPrint.Enabled = True Then
      Cancel = False
      textPrint_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textRegDate.Enabled = True Then
      Cancel = False
      textRegDate_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM14.Enabled = True Then
      Cancel = False
      textTM14_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   '2009/1/19 MODIFY BY SONIA 加CFC不控制,CFC-000769
   If m_TM01 <> "CFC" Then
      If Me.textTM21.Enabled = True Then
         Cancel = False
         textTM21_Validate Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
      
      If Me.textTM22.Enabled = True Then
         Cancel = False
         textTM22_Validate Cancel
         If Cancel = True Then
            Exit Function
         End If
         'add by sonia 2020/11/16 延展證書再檢查月日與國家檔設定不同要提醒CFT-002276
         'modify by sonia 2020/12/17 葉芳如說僅限西班牙211案
         If m_TM01 = "CFT" And m_CP10 = "102" And m_TM10 = "211" Then
            Dim strKey(0 To 4) As String, strTmp As String
            strKey(0) = m_CP09
            strKey(1) = m_TM01
            strKey(2) = m_TM02
            strKey(3) = m_TM03
            strKey(4) = m_TM04
            If TFGetMoneyDate(m_TM10, strKey, m_TM21, strTmp, New_TM22) Then
            End If
            If Mid(m_TM21, 5) <> Mid(textTM22, 5) Then
               '改回依延展檢查的模組重讀否則會影響專用期起日的檢查
               If CFTGetNewDate(m_TM10, strKey, m_TM21, New_TM22, Pre_TM22) Then
               End If
               If MsgBox("專用期止日與國家檔設定起算日不同，是否重新輸入專用期止日？", vbYesNo) = vbYes Then
                  textTM22.SetFocus
                  Exit Function
               End If
            End If
         End If
         'end 2020/11/16
      End If
   End If
   '2006/4/11 ADD BY SONIA
   Cancel = False
   Text1_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
   '2006/4/11 END
   
   'Added by Lydia 2021/08/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
   End If
   
   TxtValidate = True
End Function

'Add By Cheng 2002/06/07
Private Function GetDelayTime(strTM10 As String) As Integer
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

   StrSQLa = "Select NA15 From Nation Where NA01='" & strTM10 & "'"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      GetDelayTime = Val("0" & rsA.Fields(0).Value)
   Else
      GetDelayTime = 0
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   
End Function

