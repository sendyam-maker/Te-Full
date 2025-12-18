VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03020406_03 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標案被禁止處分"
   ClientHeight    =   5220
   ClientLeft      =   -3345
   ClientTop       =   5010
   ClientWidth     =   9135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   9135
   Begin VB.TextBox textPrint 
      Height          =   285
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   7
      Top             =   4200
      Width           =   732
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8100
      TabIndex        =   12
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5880
      TabIndex        =   10
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6840
      TabIndex        =   11
      Top             =   72
      Width           =   1212
   End
   Begin VB.TextBox textCP08 
      Height          =   285
      Left            =   1440
      MaxLength       =   40
      TabIndex        =   4
      Top             =   3480
      Width           =   2532
   End
   Begin VB.TextBox textCP26 
      Height          =   285
      Left            =   6000
      MaxLength       =   1
      TabIndex        =   8
      Top             =   4200
      Width           =   372
   End
   Begin VB.TextBox textCP48 
      Height          =   285
      Left            =   6000
      MaxLength       =   7
      TabIndex        =   6
      Top             =   3840
      Width           =   2292
   End
   Begin VB.TextBox textCP14 
      Height          =   285
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   5
      Top             =   3840
      Width           =   732
   End
   Begin VB.TextBox textRvType 
      Height          =   285
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   0
      Top             =   2040
      Width           =   372
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2292
   End
   Begin VB.TextBox textTM09 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   600
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   600
      Width           =   2532
   End
   Begin MSForms.TextBox textCP40 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   2400
      Width           =   7512
      VariousPropertyBits=   671105051
      MaxLength       =   600
      Size            =   "13250;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP41 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   2760
      Width           =   7512
      VariousPropertyBits=   671105051
      MaxLength       =   600
      Size            =   "13250;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP42 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   3120
      Width           =   7512
      VariousPropertyBits=   671105051
      MaxLength       =   600
      Size            =   "13250;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP64 
      Height          =   525
      Left            =   1440
      TabIndex        =   9
      Top             =   4560
      Width           =   7512
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13250;926"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5760
      TabIndex        =   39
      Top             =   1680
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
   Begin MSForms.TextBox textCP14_2 
      Height          =   285
      Left            =   2250
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1785
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "3149;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1170
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   1290
      Width           =   7485
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "13203;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   1170
      TabIndex        =   36
      Top             =   930
      Width           =   7485
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13203;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
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
      Left            =   3780
      TabIndex        =   35
      Top             =   630
      Width           =   645
   End
   Begin VB.Label Label21 
      Caption         =   "進度備註 :"
      Height          =   252
      Left            =   120
      TabIndex        =   34
      Top             =   4560
      Width           =   972
   End
   Begin VB.Label Label22 
      Caption         =   "列印定稿 :"
      Height          =   252
      Left            =   120
      TabIndex        =   33
      Top             =   4200
      Width           =   972
   End
   Begin VB.Label Label23 
      Caption         =   "(N:不印)"
      Height          =   252
      Left            =   2280
      TabIndex        =   32
      Top             =   4200
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "來函性質 :"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   31
      Top             =   2040
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "(1.被禁止處分 2.取消被禁止處分)"
      Height          =   252
      Index           =   2
      Left            =   2040
      TabIndex        =   30
      Top             =   2040
      Width           =   2892
   End
   Begin VB.Label Label1 
      Caption         =   "對造名稱(中) :"
      Height          =   252
      Index           =   3
      Left            =   120
      TabIndex        =   29
      Top             =   2400
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "對造名稱(英) :"
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   28
      Top             =   2760
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "對造名稱(日) :"
      Height          =   252
      Index           =   5
      Left            =   120
      TabIndex        =   27
      Top             =   3120
      Width           =   1212
   End
   Begin VB.Label Label8 
      Caption         =   "機關文號 :"
      Height          =   252
      Left            =   120
      TabIndex        =   26
      Top             =   3480
      Width           =   972
   End
   Begin VB.Label Label16 
      Caption         =   "是否算案件數 :"
      Height          =   252
      Left            =   4680
      TabIndex        =   25
      Top             =   4200
      Width           =   1212
   End
   Begin VB.Label Label15 
      Caption         =   "(N:不算)"
      Height          =   252
      Left            =   6720
      TabIndex        =   24
      Top             =   4200
      Width           =   972
   End
   Begin VB.Label Label26 
      Caption         =   "承辦期限 :"
      Height          =   252
      Left            =   4680
      TabIndex        =   23
      Top             =   3840
      Width           =   852
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   22
      Top             =   3840
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   252
      Index           =   11
      Left            =   4680
      TabIndex        =   21
      Top             =   1696
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   252
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Top             =   1680
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "商品類別 :"
      Height          =   252
      Index           =   7
      Left            =   4680
      TabIndex        =   19
      Top             =   600
      Width           =   852
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   18
      Top             =   1320
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   17
      Top             =   960
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   852
   End
End
Attribute VB_Name = "frm03020406_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/13 改成Form2.0 ;cmbTM05、textTM23、textCP13、textCP14_2、textCP64、textCP40、textCP41、textCP42
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
' 國家代碼
Dim m_TM10 As String
'
Dim m_CurrSel As Integer
'Added by Morgan 2017/5/8 電子公文
Public m_DocWord As String
Public m_DocNo As String
Public m_AppNo As String
Public m_DeadLine As String
Public m_NewCP10 As String
'end 2017/5/8

' 原資料是否有實際結果
Private Sub cmdCancel_Click()
   Unload Me
   frm03020406_02.Show
End Sub

Private Sub cmdExit_Click()
   Unload frm03020406_02
   Unload frm03020406_01
   Unload Me
End Sub

Private Sub cmdOK_Click()
   If CheckDataValid = True Then
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 儲存資料
      'edit by  nick 2004/11/03
      'OnSaveData
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      Unload Me
      Unload frm03020406_02
      'Modified by Morgan 2017/5/8 電子公文
      'frm03020406_01.Show
      If m_DocNo <> "" Then
         Unload frm03020406_01
         frm02010412.GoNext
      Else
         frm03020406_01.Show
      End If
      'end 2017/5/8
   End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM09.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   
   textCP05S.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
   
   MoveFormToCenter Me
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

' 讀取商標基本檔
Private Sub QueryTradeMark()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
      
   ' 取得商標基本檔的相關項目
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
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
      ' 商品類別
      If IsNull(rsTmp.Fields("TM09")) = False Then
         textTM09 = rsTmp.Fields("TM09")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"))
      End If
      'add by nickc 2006/05/29 加入閉卷提示
      If IsNull(rsTmp.Fields("tm29")) Then
         Me.lblClose.Caption = ""
      Else
         Me.lblClose.Caption = "已閉卷"
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 讀取案件進度檔
Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' 取得案件進度檔檔案中欄位
   strSql = "SELECT * FROM CaseProgress " & _
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
                                 "CP09 LIKE 'A%') "
            
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 總收文號
      If IsNull(rsTmp.Fields("CP09")) = False Then
         m_CP09 = rsTmp.Fields("CP09")
      End If
      ' 業務區
      If IsNull(rsTmp.Fields("CP12")) = False Then
         m_CP12 = rsTmp.Fields("CP12")
      End If
      ' 智權人員
      'Modified by Lydia 2021/08/03 改由PUB_GetFCTSalesNo帶出和產生的C類收文一致
      'If IsNull(rsTmp.Fields("CP13")) = False Then
      '   m_CP13 = rsTmp.Fields("CP13")
      '   textCP13 = GetStaffName(rsTmp.Fields("CP13"))
      'End If
      m_CP13 = Empty
      m_CP13 = PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)
      textCP13 = GetStaffName(m_CP13)
      'end 2021/08/03
      
        '預設承辦人
        Me.textCP14.Text = PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)
        Me.textCP14_2.Text = GetStaffName(Me.textCP14.Text)
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   ' 90.11.19 modify by sonia
   Dim strTmp As String
   If Len(strSrvDate(2)) = 6 Then
      strTmp = Left(strSrvDate(2), 2)
   Else
      strTmp = Left(strSrvDate(2), 3)
   End If
   If textCP08 = "" Then
      textCP08 = "（" & strTmp & "）慧商字第號"
   End If
   
   'Added by Morgan 2017/5/8 電子公文
   If m_DocWord <> "" Then
      textCP08 = m_DocWord & "字第" & PUB_GetEDocNo(m_DocNo) & "號"
   ElseIf m_DocNo <> "" Then
      textCP08 = Replace(textCP08, "第號", "第" & PUB_GetEDocNo(m_DocNo) & "號")
   End If
   '來函性質
   If m_NewCP10 = "1614" Then
      textRvType = "1"
   ElseIf m_NewCP10 = "1615" Then
      textRvType = "2"
   End If
   'end 2017/5/8
End Sub

Public Sub QueryData()
   Dim strDay As String
   ' 來函收文日
   textCP05S = m_CP05
   ' 本所案號
   textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
   
   ' 讀取商標基本檔
   QueryTradeMark
   
   ' 讀取案件進度檔
   QueryCaseProgress
   
   ' 以來函性質"被禁止處分"或"取消被禁止處"計算承辦期限
''''edit by nickc    2007/10/11 改抓有時效性的
''''   strDay = Empty
   Select Case textRvType
      Case "1":
''''         strDay = GetWorkDays(m_TM01, m_TM10, "1614")
            textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, "1614", DBDATE(m_CP05)))
      Case "2":
''''         strDay = GetWorkDays(m_TM01, m_TM10, "1615")
            textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, "1615", DBDATE(m_CP05)))
   End Select
''''   If IsEmptyText(strDay) = False Then
''''      ' 90.07.03 modify by louis (承辦期限以實際工作天數來計算)
''''      'textCP48 = TAIWANDATE(DateSerial(Val(DBYEAR(m_CP05)), Val(DBMONTH(m_CP05)), Val(DBDAY(m_CP05)) + Val(strDay)))
''''      textCP48 = TAIWANDATE(CompWorkDay(Val(strDay), DBDATE(m_CP05), 0))
''''   End If
   
   ' 非A類收文其預設為不可算案件數
   textCP26 = "N"
   
End Sub

'edit by nick 2004/11/03
'Public sub OnSaveData()
Public Function OnSaveData() As Boolean
OnSaveData = True
   Dim strSql As String
   Dim strCP09 As String
   Dim strCP10 As String
   'Dim strCP12 As String
   Dim strCP27 As String
      
 '911107 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 新增一筆資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   
   ' 案件性質為被禁止處分或取消禁止處分
   strCP10 = "1614"
   Select Case textRvType
      Case "1": strCP10 = "1614"
      Case "2": strCP10 = "1615"
   End Select
   ' 業務區別 91.8.26 MODIFY BY SONIA
   'strCP12 = GetStaffDepartment(m_CP13)
   ' 發文日為系統日
   strCP27 = DBDATE(SystemDate())
   ' 組成SQL語法
   ' 91.03.25 modify by louis (單引號)
    'Modify By Cheng 2003/04/07
    '智權人員存最近收文A類接洽記錄單的智權人員
    'Modify By Cheng 2003/09/05
'   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP20,CP26,CP32,CP40,CP41,CP42,CP64) " & _
'               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                       "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
'                       "'" & "N" & "','" & textCP26 & "','" & "N" & "','" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & textCP42 & "','" & ChgSQL(textCP64) & "')"
   strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP20,CP26,CP32,CP40,CP41,CP42,CP64) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                       "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & m_CP12 & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
                       "'" & "N" & "','" & textCP26 & "','" & "N" & "','" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & textCP42 & "','" & ChgSQL(textCP64) & "')"
   cnnConnection.Execute strSql
   
   ' 有輸入承辦人時
   If IsEmptyText(textCP14) = False Then
      strSql = "UPDATE CaseProgress SET CP14 = '" & textCP14 & "' " & _
               "WHERE CP09 = '" & strCP09 & "' "
      cnnConnection.Execute strSql
   End If
   ' 有輸入承辦期限時
   If IsEmptyText(textCP48) = False Then
      strSql = "UPDATE CaseProgress SET CP48 = " & DBDATE(textCP48) & " " & _
               "WHERE CP09 = '" & strCP09 & "' "
      cnnConnection.Execute strSql
   End If
   
   'Added by Morgan 2017/5/8 電子公文
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, strCP09, m_TM01, m_TM02, m_TM03, m_TM04, strCP10
   End If
   'end 2017/5/8
   
 '911107 nick transation
  cnnConnection.CommitTrans
     Exit Function
CheckingErr:
    MsgBox (Err.Description)
     cnnConnection.RollbackTrans
     'edit by nick 2004/11/03
     OnSaveData = False
End Function

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm03020406_03 = Nothing
End Sub

'Add By Sindy 2010/11/29
Private Sub textCP14_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

' 承辦人
Private Sub textCP14_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   
   Cancel = False
   textCP14_2 = Empty
   If IsEmptyText(textCP14) = False Then
      textCP14_2 = GetStaffName(textCP14)
      If IsEmptyText(textCP14_2) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "承辦人代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP14_GotFocus
      End If
   End If
End Sub

Private Sub textCP26_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 對造名稱(中)
Private Sub textCP40_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP40, 600) = False Then
      Cancel = True
      strTit = "資料檢核"
      strMsg = "對造名稱(中)內容長度太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP40_GotFocus
   End If
End Sub

' 對造名稱(英)
Private Sub textCP41_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP41, 600) = False Then
      Cancel = True
      strTit = "資料檢核"
      strMsg = "對造名稱(英)內容長度太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP41_GotFocus
   End If
End Sub

' 對造名稱(日)
Private Sub textCP42_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP42, 600) = False Then
      Cancel = True
      strTit = "資料檢核"
      strMsg = "對造名稱(日)內容長度太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP42_GotFocus
   End If
End Sub

' 承辦人期限
Private Sub textCP48_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP48) = False Then
      ' 檢查是否為民國日期
      If CheckIsTaiwanDate(textCP48, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的承辦期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48_GotFocus
      End If
   End If
End Sub

' 是否算案件數
Private Sub textCP26_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCP26) = False Then
      Select Case textCP26
         Case " ", "N":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP26_GotFocus
      End Select
   End If
End Sub

' 進度備註
Private Sub textCP64_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If CheckLengthIsOK(textCP64, 2000) = False Then
      Cancel = True
      strTit = "資料檢核"
      strMsg = "進度備註資料內容長度太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP64_GotFocus
   End If
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'來函性質
Private Sub textRvType_Validate(Cancel As Boolean)
   Dim strDay As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textRvType) = False Then
      Select Case textRvType
         Case "1", "2":
         Case Else:
            Cancel = True
            strTit = "檢核資料"
            strMsg = "來函性質只可輸入1或2"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textRvType_GotFocus
      End Select
      
      ' 以案件性質"審查報告"或"核駁前先行通知"計算承辦期限
''''edit by nickc 2007/10/11 改抓有時效性的
''''      strDay = Empty
      Select Case textRvType
         Case "1":
''''            strDay = GetWorkDays(m_TM01, m_TM10, "1614")
                textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, "1614", DBDATE(m_CP05)))
         Case "2":
''''            strDay = GetWorkDays(m_TM01, m_TM10, "1615")
                textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, "1615", DBDATE(m_CP05)))
      End Select
''''      If IsEmptyText(strDay) = False Then
''''         ' 90.07.03 modify by louis (承辦期限以實際工作天數來計算)
''''         'textCP48 = TAIWANDATE(DateSerial(Val(DBYEAR(m_CP05)), Val(DBMONTH(m_CP05)), Val(DBDAY(m_CP05)) + Val(strDay)))
''''         textCP48 = TAIWANDATE(CompWorkDay(Val(strDay), DBDATE(m_CP05), 0))
''''      End If
   End If
End Sub

' 列印定稿
Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPrint) = False Then
      Select Case textPrint
         Case " ", "N":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrint_GotFocus
      End Select
   End If
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   
   ' 機關文號不可以空白
   If IsEmptyText(textCP08) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入機關文號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP08.SetFocus
      GoTo EXITSUB
   End If
   ' 來函性質不可以空白
   If IsEmptyText(textRvType) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入來函性質"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textRvType.SetFocus
      GoTo EXITSUB
   End If
   ' 承辦期限不可以空白
   If IsEmptyText(textCP48) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入承辦期限"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP48.SetFocus
      GoTo EXITSUB
   End If
   ' 對造中英日文名稱不可同時空白
   If IsEmptyText(textCP40) = True And IsEmptyText(textCP41) = True And IsEmptyText(textCP42) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入對造名稱"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP40.SetFocus
      GoTo EXITSUB
   End If
   
   'Added by Lydia 2021/09/13 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textRvType_GotFocus()
   InverseTextBox textRvType
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textCP08_GotFocus()
   'Modify By Cheng 2002/04/22
   '將游標停在"字"的前面
'   InverseTextBox textCP08
Dim intPos As Integer
With Me.textCP08
   If Len("" & .Text) > 0 Then
      intPos = InStr("" & .Text, "字")
      If intPos - 1 >= 0 Then
         .SelStart = intPos - 1
         .SelLength = 0
      End If
   End If
End With
End Sub

Private Sub textCP14_GotFocus()
   InverseTextBox textCP14
End Sub

Private Sub textCP26_GotFocus()
   InverseTextBox textCP26
End Sub

Private Sub textCP40_GotFocus()
   InverseTextBox textCP40
End Sub

Private Sub textCP41_GotFocus()
   InverseTextBox textCP41
End Sub

Private Sub textCP42_GotFocus()
   InverseTextBox textCP42
End Sub

Private Sub textCP48_GotFocus()
   InverseTextBox textCP48
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

