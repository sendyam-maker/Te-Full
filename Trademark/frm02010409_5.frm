VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010409_5 
   BorderStyle     =   1  '單線固定
   Caption         =   "服務業務結果輸入(條碼)"
   ClientHeight    =   5470
   ClientLeft      =   200
   ClientTop       =   1010
   ClientWidth     =   9130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5470
   ScaleWidth      =   9130
   Begin VB.TextBox textModLetter 
      Height          =   264
      Left            =   6840
      MaxLength       =   3
      TabIndex        =   7
      Top             =   4680
      Width           =   372
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "測試是否合格(&T)"
      Height          =   400
      Left            =   4656
      TabIndex        =   9
      Top             =   70
      Width           =   1500
   End
   Begin VB.TextBox textPeriod 
      Height          =   264
      Left            =   4860
      MaxLength       =   3
      TabIndex        =   5
      Top             =   4320
      Width           =   612
   End
   Begin VB.TextBox textCP16 
      Height          =   264
      Left            =   5940
      MaxLength       =   8
      TabIndex        =   2
      Top             =   3960
      Width           =   1092
   End
   Begin VB.TextBox textSP19 
      Height          =   264
      Left            =   1440
      MaxLength       =   7
      TabIndex        =   1
      Top             =   3960
      Width           =   1092
   End
   Begin VB.TextBox textCP08 
      Height          =   264
      Left            =   1440
      MaxLength       =   40
      TabIndex        =   0
      Top             =   3600
      Width           =   2532
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7008
      TabIndex        =   11
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6180
      TabIndex        =   10
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8232
      TabIndex        =   12
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox textPrint 
      Height          =   264
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   6
      Top             =   4680
      Width           =   732
   End
   Begin VB.TextBox textSP20 
      Height          =   264
      Left            =   1440
      MaxLength       =   7
      TabIndex        =   3
      Top             =   4320
      Width           =   1092
   End
   Begin VB.TextBox textSP21 
      Height          =   264
      Left            =   2880
      MaxLength       =   7
      TabIndex        =   4
      Top             =   4320
      Width           =   1092
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3240
      Width           =   2532
   End
   Begin VB.TextBox textSPKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   720
      Width           =   2532
   End
   Begin VB.TextBox textSP06 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1440
      Width           =   7572
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2880
      Width           =   2532
   End
   Begin VB.TextBox textCP05 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5580
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2520
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2520
      Width           =   2532
   End
   Begin MSForms.TextBox textCP64 
      Height          =   300
      Left            =   1440
      TabIndex        =   8
      Top             =   5040
      Width           =   7492
      VariousPropertyBits=   -1467989989
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13215;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   264
      Left            =   5580
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2880
      Width           =   2532
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textSP08 
      Height          =   264
      Left            =   1440
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2160
      Width           =   7572
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "13356;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textSP05 
      Height          =   264
      Left            =   1440
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1080
      Width           =   7572
      VariousPropertyBits=   679493663
      Size            =   "13356;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textSP07 
      Height          =   264
      Left            =   1440
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1800
      Width           =   7572
      VariousPropertyBits=   679493663
      Size            =   "13356;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "Y:修改"
      Height          =   255
      Index           =   12
      Left            =   7320
      TabIndex        =   43
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "是否修改定稿內容 :"
      Height          =   255
      Left            =   5040
      TabIndex        =   42
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "期登記期"
      Height          =   252
      Index           =   9
      Left            =   5580
      TabIndex        =   41
      Top             =   4320
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "第"
      Height          =   252
      Index           =   8
      Left            =   4620
      TabIndex        =   40
      Top             =   4320
      Width           =   252
   End
   Begin VB.Label Label1 
      Caption         =   "製作正片費用 :"
      Height          =   252
      Index           =   7
      Left            =   4620
      TabIndex        =   39
      Top             =   3960
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "廠商號碼 :"
      Height          =   252
      Index           =   5
      Left            =   120
      TabIndex        =   38
      Top             =   3960
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "機關文號 :"
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   37
      Top             =   3600
      Width           =   1092
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "(N:不印;1:台->各國;2:外->台;3:英文)"
      Height          =   180
      Left            =   2220
      TabIndex        =   36
      Top             =   4740
      Width           =   2745
   End
   Begin VB.Label Label22 
      Caption         =   "列印定稿 :"
      Height          =   252
      Left            =   120
      TabIndex        =   35
      Top             =   4680
      Width           =   972
   End
   Begin VB.Label Label21 
      Caption         =   "進度備註 :"
      Height          =   252
      Left            =   120
      TabIndex        =   34
      Top             =   5040
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "使用期間 :"
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   33
      Top             =   4320
      Width           =   1092
   End
   Begin VB.Line Line1 
      X1              =   2640
      X2              =   2760
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   252
      Index           =   11
      Left            =   4620
      TabIndex        =   32
      Top             =   2880
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   252
      Index           =   10
      Left            =   120
      TabIndex        =   31
      Top             =   3240
      Width           =   1092
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   30
      Top             =   2160
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   29
      Top             =   720
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "案件中文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   28
      Top             =   1080
      Width           =   1332
   End
   Begin VB.Label Label9 
      Caption         =   "案件英文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   27
      Top             =   1440
      Width           =   1212
   End
   Begin VB.Label Label10 
      Caption         =   "案件日文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   26
      Top             =   1800
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   252
      Index           =   6
      Left            =   120
      TabIndex        =   25
      Top             =   2880
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "收文日 :"
      Height          =   252
      Index           =   3
      Left            =   4620
      TabIndex        =   24
      Top             =   2520
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   23
      Top             =   2520
      Width           =   732
   End
End
Attribute VB_Name = "frm02010409_5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/29 Form2.0已修改 textSP05/textSP07/textSP08/textCP13/textCP64
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
Option Explicit

' 本所案號
Dim m_SP01 As String
Dim m_SP02 As String
Dim m_SP03 As String
Dim m_SP04 As String
' 申請國家
Dim m_SP09 As String
' 來函收文日
Dim m_CP05 As String
' 所選取的收文號
Dim m_CP09 As String
' 案件性質
Dim m_CP10 As String
' 智權人員
Dim m_CP13 As String
Dim m_CP12 As String
' 申請人
Dim m_TM23 As String
'Add By Cheng 2002/11/27
'原承辦人
Dim m_CP14 As String
'Add By Cheng 2003/07/16
Dim m_SP22 As String '正片號碼
'Add By Sindy 2019/5/22
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2019/5/22 END
Dim strLD18 As String 'Add By Sindy 2019/12/20 信函總收文號
Dim m_TM44 As String 'Add By Sindy 2019/12/20 FC代理人


'Add By Sindy 2019/5/22
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdCancel_Click()
   frm02010409_2.Show
   Unload Me
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm02010409_2
   Unload frm02010409_1
   Unload Me
End Sub

Private Sub cmdok_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If CheckDataValid() = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
          'add by nickc 2005/04/22
          Pub_EndModCashMsg m_SP09
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 儲存資料
        'Modify By Cheng 2002/11/07
'      'OnSaveData
        If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
        'Modify By Cheng 2002/11/08
        ' 列印定稿
        If textPrint <> "N" Then
           PrintLetter
        End If
      
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      Unload frm02010409_2
      'Add By Sindy 2019/5/22
      If Me.m_strIR01 <> "" Then
         Unload frm02010409_1
         If Not m_PrevForm Is Nothing Then
            Call m_PrevForm.GoNext
         End If
      '2019/5/22 END
      Else
         frm02010409_1.Show
      End If
      Unload Me
   End If
End Sub

Private Sub cmdTest_Click()
   frm02010409_9.SetData 0, m_SP01, True
   frm02010409_9.SetData 1, m_SP02, False
   frm02010409_9.SetData 2, m_SP03, False
   frm02010409_9.SetData 3, m_SP04, False
   frm02010409_9.Show
   frm02010409_9.QueryData
   Me.Hide
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textSPKey.BackColor = &H8000000F
   textSP05.BackColor = &H8000000F
   textSP06.BackColor = &H8000000F
   textSP07.BackColor = &H8000000F
   textSP08.BackColor = &H8000000F
   textCP05.BackColor = &H8000000F
   textCP05S.BackColor = &H8000000F
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   
   MoveFormToCenter Me
   
   'Add By Sindy 2019/5/22
   m_strIR01 = frm02010409_1.m_strIR01
   m_strIR02 = frm02010409_1.m_strIR02
   m_strIR03 = frm02010409_1.m_strIR03
   m_strIR04 = frm02010409_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2019/5/22 END
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_SP01 = Empty
      m_SP02 = Empty
      m_SP03 = Empty
      m_SP04 = Empty
      m_CP05 = Empty
   End If
   
   Select Case nType
      ' 本所案號 欄位1
      Case 0: m_SP01 = strData
      ' 本所案號 欄位2
      Case 1: m_SP02 = strData
      ' 本所案號 欄位3
      Case 2: m_SP03 = strData
      ' 本所案號 欄位4
      Case 3: m_SP04 = strData
      ' 來函收文日
      Case 4: m_CP05 = strData
      ' 收文號
      Case 5: m_CP09 = strData
   End Select
End Sub

' 取得服務業務基本檔
Private Sub QueryServicePractice()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   ' 取得商標基本檔的相關項目
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_SP01 & "' AND " & _
                  "SP02 = '" & m_SP02 & "' AND " & _
                  "SP03 = '" & m_SP03 & "' AND " & _
                  "SP04 = '" & m_SP04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請國家
      'Add By Cheng 2002/07/17
      m_SP09 = Empty
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_SP09 = rsTmp.Fields("SP09")
      End If
      ' 商標名稱(中)
      If IsNull(rsTmp.Fields("SP05")) = False Then
         textSP05 = rsTmp.Fields("SP05")
      End If
      ' 商標名稱(英)
      If IsNull(rsTmp.Fields("SP06")) = False Then
         textSP06 = rsTmp.Fields("SP06")
      End If
      ' 商標名稱(日)
      If IsNull(rsTmp.Fields("SP07")) = False Then
         textSP07 = rsTmp.Fields("SP07")
      End If
      ' 申請人
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("SP08")) = False Then
         m_TM23 = rsTmp.Fields("SP08")
         textSP08 = GetCustomerName(rsTmp.Fields("SP08"), 0)
      End If
      
      'Add By Sindy 2019/12/20
      ' FC代理人
      m_TM44 = Empty
      If IsNull(rsTmp.Fields("SP26")) = False Then
         m_TM44 = rsTmp.Fields("SP26")
      End If
      '2019/12/20 END
            
      ' 廠商號碼
      If IsNull(rsTmp.Fields("SP19")) = False Then
         textSP19 = rsTmp.Fields("SP19")
      End If
      ' 使用期間 (起)
      If IsNull(rsTmp.Fields("SP20")) = False Then
         textSP20 = TAIWANDATE(rsTmp.Fields("SP20"))
      End If
      ' 使用期間 (迄)
      If IsNull(rsTmp.Fields("SP21")) = False Then
         textSP21 = TAIWANDATE(rsTmp.Fields("SP21"))
      End If
        'Add By Cheng 2003/07/16
        '正片號碼
        m_SP22 = "" & rsTmp.Fields("SP22")
        'add by nickc 2006/11/21
        textPrint = CheckStr(rsTmp.Fields("SP72"))
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 查詢資料庫取得資料
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   m_CP13 = Empty
   m_CP12 = Empty
   'Add By Cheng 2002/11/27
   m_CP14 = Empty
   
   ' 來函收文日
   textCP05S = m_CP05
   ' 收文號
   textCP09 = m_CP09
   ' 讀取服務業務基本檔檔案
   QueryServicePractice
   
   ' 本所案號
   textSPKey = m_SP01 & m_SP02 & m_SP03 & m_SP04
   
   ' 取得案件進度檔A類資料的最後一筆
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP01 = '" & m_SP01 & "' AND " & _
                  "CP02 = '" & m_SP02 & "' AND " & _
                  "CP03 = '" & m_SP03 & "' AND " & _
                  "CP04 = '" & m_SP04 & "' AND " & _
                  "CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 收文日
      If IsNull(rsTmp.Fields("CP05")) = False Then
         textCP05 = TAIWANDATE(rsTmp.Fields("CP05"))
      End If
      ' 收文號
      If IsNull(rsTmp.Fields("CP09")) = False Then
         textCP09 = rsTmp.Fields("CP09")
      End If
      ' 案件性質
      m_CP10 = Empty
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_SP09 < "010" Then
            textCP10 = GetCaseTypeName(m_SP01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_SP01, rsTmp.Fields("CP10"), 1)
         End If
         'Added by Lydia 2017/08/21
         If m_SP01 = "TB" And m_CP10 = "708" Then
            textPeriod = "" & rsTmp.Fields("CP53") 'Added by Lydia 2017/08/21 第?期登記期'欄預設該筆之CP53
         End If
         'end 2017/08/21
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"), True)
      End If
      '業務區   nick 91.08.22
      If IsNull(rsTmp.Fields("cp12")) = False Then
        m_CP12 = rsTmp.Fields("cp12")
      End If
    'Add By Cheng 2002/11/27
    m_CP14 = "" & rsTmp("CP14").Value
   End If
   rsTmp.Close
   
   ' 當案件性質為正片製作時才可輸入製作正片費用
   If m_CP10 = "803" Then
      EnableTextBox textCP16, True
   Else
      EnableTextBox textCP16, False
   End If
   ' 當案件性質為繳年費時, 第幾期登記期才可輸入
   If m_CP10 = "708" Then
      EnableTextBox textPeriod, True
   Else
      EnableTextBox textPeriod, False
   End If
   ' 當案件性質為條碼申請或繳年費時, 使用期間才可輸入
   If m_CP10 = "802" Or m_CP10 = "708" Then
      EnableTextBox textSP20, True
      EnableTextBox textSP21, True
   Else
      EnableTextBox textSP20, False
      EnableTextBox textSP21, False
   End If
   
   Set rsTmp = Nothing
   
   'add by nickc 2006/06/30 帶列印定稿預設值
   'edit by nickc 2006/11/21
   If textPrint = "" Then
        textPrint = GetTWordLng(m_SP01, m_SP02, m_SP03, m_SP04)
   End If
End Sub

' 儲存資料
'Modify By Cheng 2002/11/07
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strCP09 As String
   Dim strCP10 As String
   Dim strCP12 As String
   Dim strCP16 As String
   Dim strCP20 As String
   Dim strCP27 As String
   Dim strCP32 As String
   Dim strSP20 As String
   Dim strSP21 As String
   Dim strNP07 As String
   Dim strNP08 As String
   Dim strNP09 As String
   Dim strNP22 As String
   
'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
OnSaveData = True
cnnConnection.BeginTrans
   
   ' 新增一筆資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   ' 案件性質為服務業務結果
   strCP10 = "1801"
   ' 業務區別
   'strCP12 = GetST15(m_CP13)
   ' 費用
   strCP16 = Empty
   strCP20 = "N"
   strCP32 = "N"
   If m_CP10 = "803" And IsEmptyText(textCP16) = False Then
      strCP16 = textCP16
      strCP20 = "null"
      strCP32 = "null"
   End If
   ' 發文日為系統日
   strCP27 = DBDATE(SystemDate())
   If strCP16 = Empty Then
      ' 91.03.25 modify by louis (單引號)
        '若案件性質為正片測試(804)
        If m_CP10 = "804" Then
            '承辦人為原程序承辦人, 不上發文日
            'Modify By Cheng 2003/04/03
            '智權人員存最近收文A類接洽記錄單的智權人員
            'Modify By Cheng 2004/02/03
            '業務區為最近收文A類接洽記錄單智權人員的業務區
'            strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP26,CP43,CP64) " & _
'                     "VALUES ('" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & DBDATE(m_CP05) & "," & _
'                             "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "','" & m_CP14 & "'," & _
'                             "'" & "N" & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "')"
            strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP26,CP43,CP64) " & _
                     "VALUES ('" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & DBDATE(m_CP05) & "," & _
                             "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04)) & "','" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "','" & m_CP14 & "'," & _
                             "'" & "N" & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "')"
            'End
        Else
            '承辦人為使用者, 發文日為系統日
            'Modify By Cheng 2003/04/03
            '智權人員存最近收文A類接洽記錄單的智權人員
            'Modify By Cheng 2004/02/03
            '業務區為最近收文A類接洽記錄單智權人員的業務區
'            strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP26,CP27,CP43,CP64) " & _
'                     "VALUES ('" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & DBDATE(m_CP05) & "," & _
'                             "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "','" & strUserNum & "'," & _
'                             "'" & "N" & "'," & strCP27 & ",'" & m_CP09 & "','" & ChgSQL(textCP64) & "')"
            '2015/1/14 modify by sonia 所有服務業務結果的承辦人改放原承辦人TM-000067(宋若蘭),否則期限表帶出之承辦人會是程序
            'strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP26,CP27,CP43,CP64) " & _
                     "VALUES ('" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & DBDATE(m_CP05) & "," & _
                             "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04)) & "','" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "','" & strUserNum & "'," & _
                             "'" & "N" & "'," & strCP27 & ",'" & m_CP09 & "','" & ChgSQL(textCP64) & "')"
            strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP26,CP27,CP43,CP64) " & _
                     "VALUES ('" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & DBDATE(m_CP05) & "," & _
                             "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04)) & "','" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "','" & m_CP14 & "'," & _
                             "'" & "N" & "'," & strCP27 & ",'" & m_CP09 & "','" & ChgSQL(textCP64) & "')"
            'End
        End If
   Else
      ' 91.03.25 modify by louis (單引號)
      'strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP17,CP20,CP26,CP27,CP32,CP43,CP64) " & _
               "VALUES ('" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & DBDATE(m_CP05) & "," & _
                       "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & m_CP13 & "','" & strUserNum & "'," & _
                       "" & strCP16 & ",'" & strCP20 & "','" & "N" & "'," & strCP27 & ",'" & strCP32 & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "')"
      'nick 911018
        '若案件性質為正片測試(804)
        If m_CP10 = "804" Then
            '承辦人為原程序承辦人, 不上發文日
            'Modify By Cheng 2003/04/03
            '智權人員存最近收文A類接洽記錄單的智權人員
            'Modify By Cheng 2004/02/03
            '業務區為最近收文A類接洽記錄單智權人員的業務區
'            strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP17,CP20,CP26,CP32,CP43,CP64,cp18) " & _
'                     "VALUES ('" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & DBDATE(m_CP05) & "," & _
'                             "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "','" & m_CP14 & "'," & _
'                             "" & strCP16 & "," & CNULL(strCP20) & ",'" & "N" & "','" & CNULL(strCP32) & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "'," & strCP18 & ")"
            strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP16,CP20,CP26,CP32,CP43,CP64,CP17,cp18) " & _
                     "VALUES ('" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & DBDATE(m_CP05) & "," & _
                             "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04)) & "','" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "','" & m_CP14 & "'," & _
                             "" & strCP16 & "," & CNULL(strCP20) & ",'" & "N" & "','" & CNULL(strCP32) & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "'," & strCP16 & ",0)"
            'End
        Else
            '承辦人為使用者, 發文日為系統日
            'Modify By Cheng 2003/04/03
            '智權人員存最近收文A類接洽記錄單的智權人員
            'Modify By Cheng 2004/02/03
            '業務區為最近收文A類接洽記錄單智權人員的業務區
'            strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP17,CP20,CP26,CP27,CP32,CP43,CP64,cp18) " & _
'                     "VALUES ('" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & DBDATE(m_CP05) & "," & _
'                             "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "','" & strUserNum & "'," & _
'                             "" & strCP16 & "," & CNULL(strCP20) & ",'" & "N" & "'," & strCP27 & ",'" & CNULL(strCP32) & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "'," & strCP18 & ")"
            '2015/1/14 modify by sonia 所有服務業務結果的承辦人改放原承辦人TM-000067(宋若蘭),否則期限表帶出之承辦人會是程序
            'strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP16,CP20,CP26,CP27,CP32,CP43,CP64,CP17,cp18) " & _
                     "VALUES ('" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & DBDATE(m_CP05) & "," & _
                             "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04)) & "','" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "','" & strUserNum & "'," & _
                             "" & strCP16 & "," & strCP20 & ",'" & "N" & "'," & strCP27 & "," & strCP32 & ",'" & m_CP09 & "','" & ChgSQL(textCP64) & "'," & strCP16 & ",0)"
            strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP16,CP20,CP26,CP27,CP32,CP43,CP64,CP17,cp18) " & _
                     "VALUES ('" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & DBDATE(m_CP05) & "," & _
                             "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04)) & "','" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "','" & m_CP14 & "'," & _
                             "" & strCP16 & "," & strCP20 & ",'" & "N" & "'," & strCP27 & "," & strCP32 & ",'" & m_CP09 & "','" & ChgSQL(textCP64) & "'," & strCP16 & ",0)"
            'End
        End If
   End If
   cnnConnection.Execute strSql
                    
   'Add By Sindy 2019/12/20 商標電子化
   If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
      strLD18 = strCP09
      PUB_AddLetterProgress strLD18, 1, IIf(textPrint = "N", False, True), "", False, m_TM23, strCP10, m_TM44
   End If
   '2019/12/20 END
   
    'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
    Pub_UpdateFromMaxCP27 m_SP01, m_SP02, m_SP03, m_SP04
                    
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 當案件性質為條碼申請或繳年費時, 更新服務業務基本檔的使用期間欄位
   If m_CP10 = "802" Or m_CP10 = "708" Then
      strSql = "UPDATE ServicePractice SET SP20 = " & DBDATE(textSP20) & ", " & _
                                          "SP21 = " & DBDATE(textSP21) & " " & _
               "WHERE SP01 = '" & m_SP01 & "' AND " & _
                     "SP02 = '" & m_SP02 & "' AND " & _
                     "SP03 = '" & m_SP03 & "' AND " & _
                     "SP04 = '" & m_SP04 & "' "
      cnnConnection.Execute strSql
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新廠商號碼欄位
   strSql = "UPDATE ServicePractice SET SP19 = '" & textSP19 & "' " & _
            "WHERE SP01 = '" & m_SP01 & "' AND " & _
                  "SP02 = '" & m_SP02 & "' AND " & _
                  "SP03 = '" & m_SP03 & "' AND " & _
                  "SP04 = '" & m_SP04 & "' "
   cnnConnection.Execute strSql
                  
   'add by nickc 2006/11/21
   If textPrint <> "N" Then
        strSql = "UPDATE ServicePractice SET SP72 = '" & textPrint & "' " & _
                 "WHERE SP01 = '" & m_SP01 & "' AND " & _
                       "SP02 = '" & m_SP02 & "' AND " & _
                       "SP03 = '" & m_SP03 & "' AND " & _
                       "SP04 = '" & m_SP04 & "' "
        cnnConnection.Execute strSql
   End If
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新案件進度檔所選取收文資料的實際結果為 1
   strSql = "UPDATE CaseProgress SET CP24 = '1' " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   cnnConnection.Execute strSql
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 當案件性質為條碼申請或繳年費時, 新增資料到下一程序檔
   If m_CP10 = "802" Or m_CP10 = "708" Then
      ' 下一程序為繳年費
      strNP07 = "708"
      ' 法定期限為專用期限截止日
      strNP09 = DBDATE(textSP21)
      ' 本所期限為法定期限-2天
        'Modify By Cheng 2003/09/01
'      strNP08 = DBDATE(Format(DateSerial(Val(DBYEAR(strNP09)), Val(DBMONTH(strNP09)), Val(DBDAY(strNP09)) - 2)))
      'Modify By Sindy 2014/10/6 台灣案之本所期限設定
      If m_SP09 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
         strNP08 = PUB_GetOurDeadline(DBDATE(strNP09))
      Else
      '2014/10/6 END
         strNP08 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(strNP09))))
      End If
      strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      ' 序號
      strNP22 = GetNextProgressNo()
        'Modify By Cheng 2002/11/28
        If m_CP10 = "802" Then
            ' SQL 語法
            'Modify By Cheng 2003/04/03
            '智權人員存最近收文A類接洽記錄單的智權人員
            strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                     "VALUES ('" & strCP09 & "','" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & _
                             "'" & strNP07 & "'," & strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "'," & strNP22 & ")"
        Else
            'Modify By Cheng 2002/11/28
            '改成更新原B類收文號資料
            strSql = " Update NextProgress Set NP01 ='" & strCP09 & "' , NP08=" & strNP08 & ", NP09 =" & strNP09 & " Where NP01='" & m_CP09 & "' And NP07='" & m_CP10 & "'"
        End If
        cnnConnection.Execute strSql
      ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
'      '92.6.8 SONIA 加 言詞辯論, 準備程序
      Select Case strNP07
'         Case "102", "105", "702", "708", "305", "998", "997", "204", "205":
         Case "102", "105", "702", "708", "305", "998", "997"
         Case Else:
            ' 列印國內案件接洽及結案記錄單
'            g_PrtForm001.PrintForm strNP22, m_SP01, m_SP02, m_SP03, m_SP04
            'Add By Cheng 2004/04/08
            '新增列印接洽結案單資料
            pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_SP01, "" & m_SP02, "" & m_SP03, "" & m_SP04
      End Select
   End If
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Modify By Cheng 2002/11/08
'   ' 列印定稿
'   If textPrint <> "N" Then
'      PrintLetter
'   End If
          'add by nickc 2005/04/22
          Pub_UpdateEndModCash m_SP01, m_SP02, m_SP03, m_SP04
   
   'Add by Sindy 2019/5/22
   Call PUB_TMFilePathToCPP(strTMCppFilePath, strCP09) '檢查是否有電子檔要存入卷宗區
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm02010409_1"
   End If
   '2019/5/22 END
   
'Add By Cheng 2002/11/07
cnnConnection.CommitTrans
Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    OnSaveData = False
End Function

Private Sub Form_Unload(Cancel As Integer)
'edit by nickc 2008/04/25 改整批印
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   
   'Add By Sindy 2019/5/22
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   'Add By Cheng 2002/07/18
   Set frm02010409_5 = Nothing
End Sub

' 進度備註
Private Sub textCP64_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP64, 2000) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "進度備註欄位內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP64_GotFocus
   End If
End Sub

Private Sub textModLetter_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否修改定稿內容
Private Sub textModLetter_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textModLetter) = False Then
      Select Case textModLetter
         Case "", " ", "Y":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "請輸入Y或空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textModLetter_GotFocus
      End Select
   End If
End Sub

' 第#期登記期
Private Sub textPeriod_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textPeriod) = False Then
      If IsNumeric(textPeriod) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的數值"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPeriod_GotFocus
      End If
   End If
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'add by nickc 2006/06/29
   If KeyAscii <> 78 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 8 And KeyAscii <> 13 Then
       KeyAscii = 0
   End If
End Sub

' 檢查是否列印定稿欄位
Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPrint) = False Then
      Select Case textPrint
         'edit by nickc 2006/06/29
         'Case " ", "N":
         Case "N", "1", "2", "3":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            'edit by nickc 2006/06/29
            'strMsg = "只可輸入空白或N"
            strMsg = "只可輸入 N 或 1-3"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrint_GotFocus
      End Select
   End If
End Sub

' 檢查該輸入的資料是否已完成
Private Function CheckDataValid()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   'Add by Amy 2021/12/29檢查畫面的 TextBox是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        GoTo EXITSUB
   End If

   ' 廠商號碼不可空白
   If IsEmptyText(textSP19) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入廠商號碼"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textSP19.SetFocus
      GoTo EXITSUB
   End If
   
   ' 使用期間
   If m_CP10 = "802" Or m_CP10 = "708" Then
      If IsEmptyText(textSP20) = True Or IsEmptyText(textSP21) = True Then
         strTit = "資料檢核"
         strMsg = "使用期間不可空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP20.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   ' 繳年費
   If m_CP10 = "708" Then
      If IsEmptyText(textPeriod) = True Then
         strTit = "資料檢核"
         strMsg = "案件性質為繳年費, 請輸入第?期登記期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPeriod.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   '2005/12/26 MODIFY BY SONIA 可不輸入費用
   ' 案件性質為製作正片時, 製做正片費用不可空白
   If m_CP10 = "803" Then
      If IsEmptyText(textCP16) = True Then
         strTit = "資料檢核"
         strMsg = "案件性質為製作正片, 是否確定不輸入製做正片費用？"
         '2005/12/26 MODIFY BY SONIA
         'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         'textCP16.SetFocus
         'GoTo EXITSUB
         nResponse = MsgBox(strMsg, vbOKCancel, strTit)
         If nResponse <> 1 Then
            textCP16.SetFocus
            GoTo EXITSUB
         End If
      End If
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

' 使用期間(起)
Private Sub textSP20_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textSP20) = False Then
      ' 檢查是否為民國日期
      If CheckIsTaiwanDate(textSP20, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的使用期間"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP20_GotFocus
      End If
   End If
End Sub

' 使用期間(迄)
Private Sub textSP21_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textSP21) = False Then
      ' 檢查是否為民國日期
      If CheckIsTaiwanDate(textSP21, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的使用期間"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP21_GotFocus
      End If
   End If
End Sub

Private Sub textPeriod_GotFocus()
   InverseTextBox textPeriod
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textModLetter_GotFocus()
   InverseTextBox textModLetter
End Sub

Private Sub textCP08_GotFocus()
   InverseTextBox textCP08
End Sub

Private Sub textSP19_GotFocus()
   InverseTextBox textSP19
End Sub

Private Sub textSP20_GotFocus()
   InverseTextBox textSP20
End Sub

Private Sub textSP21_GotFocus()
   InverseTextBox textSP21
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
Dim strTM23Nation As String
Dim strSql As String
Dim strTmp As String
Dim ii As Integer
Dim arrSP22
   
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
      
   ' 系統別為TB
   If m_SP01 = "TB" Then
      Select Case m_CP10
         ' 條碼申請
         Case "802":
            'add by nickc 2006/06/30
            If textPrint = "1" Then
                ' 清除定稿例外欄位檔原有資料
                EndLetter "06", m_CP09, "03", strUserNum
            End If
         ' 製作正片
         Case "803":
            'add by nickc 2006/06/30
            If textPrint = "1" Then
                ' 清除定稿例外欄位檔原有資料
                EndLetter "06", m_CP09, "04", strUserNum
                ' 製作正片費用
                If textCP16 <> "" Then
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "06" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & "'," & _
                            "'" & "製作正片費用" & "','" & Chr(13) & "　　此次正片製作費用共新台幣" & textCP16 & "元整，檢附收據正本乙紙，請速惠予將該筆款項擲寄本所。" & "')"
                   cnnConnection.Execute strSql
                End If
            End If
         ' 測試正片
         Case "804":
            'add by nickc 2006/06/30
            If textPrint = "1" Then
                ' 清除定稿例外欄位檔原有資料
                EndLetter "06", m_CP09, "05", strUserNum
                'Add By Cheng 2003/07/16
                If m_SP22 <> "" Then
                    arrSP22 = Split(m_SP22, ",")
                    m_SP22 = ""
                    For ii = 0 To UBound(arrSP22)
                        If Trim(arrSP22(ii)) <> "" Then
                            m_SP22 = m_SP22 & arrSP22(ii) & "、"
                        End If
                    Next ii
                    If m_SP22 <> "" Then m_SP22 = Left(m_SP22, Len(m_SP22) - 1)
                    ' 正片號碼
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "06" & "','" & m_CP09 & "','" & "05" & "','" & strUserNum & "'," & _
                             "'" & "正片號碼" & "','" & m_SP22 & "')"
                    cnnConnection.Execute strSql
                End If
            End If
         ' 移轉
         Case "501":
            'add by nickc 2006/06/30
            If textPrint = "1" Then
                ' 清除定稿例外欄位檔原有資料
                'Modify By Cheng 2003/01/28
                '修改處理狀況
    '            EndLetter "06", m_CP09, "06", strUserNum
                EndLetter "06", m_CP09, "04", strUserNum
            End If
         ' 繳年費
         Case "708":
            'add by nickc 2006/06/30
            If textPrint = "1" Then
                ' 清除定稿例外欄位檔原有資料
                EndLetter "06", m_CP09, "07", strUserNum
                ' 條碼年費年度
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "06" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & "'," & _
                         "'" & "條碼年費年度" & "','" & textPeriod & "')"
                cnnConnection.Execute strSql
            End If
      End Select
   End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
Dim strTM23Nation As String
'Add By Sindy 2012/1/13
Dim ET01 As String, ET02 As String, ET03 As String, bolEdit As Boolean
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
'2012/1/13 End
   
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   'Add By Sindy 2012/1/13
   ET01 = "06"
   ET02 = m_CP09
   bolEdit = IIf(Me.textModLetter.Text = "Y", True, False)
   '2012/1/13 End
   
   ' 系統別為TB
   If m_SP01 = "TB" Then
      Select Case m_CP10
         ' 條碼申請
         Case "802":
            'add by nickc 2006/06/30
            If textPrint = "1" Then
                ' 列印定稿
'                NowPrint m_CP09, "06", "03", IIf(Me.textModLetter.Text = "Y", True, False), strUserNum, 0
                ET03 = "03" 'Modify By Sindy 2012/1/13
            End If
         ' 製作正片
         Case "803":
            If textModLetter = "Y" Then
                'add by nickc 2006/06/30
                If textPrint = "1" Then
                   ' 列印定稿
'                   NowPrint m_CP09, "06", "04", IIf(Me.textModLetter.Text = "Y", True, False), strUserNum, 0
                  ET03 = "04" 'Modify By Sindy 2012/1/13
                End If
            Else
                'add by nickc 2006/06/30
                If textPrint = "1" Then
                   ' 列印定稿
'                   NowPrint m_CP09, "06", "04", IIf(Me.textModLetter.Text = "Y", True, False), strUserNum, 0
                  ET03 = "04" 'Modify By Sindy 2012/1/13
                End If
            End If
         ' 測試正片
         Case "804":
            'add by nickc 2006/06/30
            If textPrint = "1" Then
                ' 列印定稿
'                NowPrint m_CP09, "06", "05", IIf(Me.textModLetter.Text = "Y", True, False), strUserNum, 0
               ET03 = "05" 'Modify By Sindy 2012/1/13
            End If
         ' 移轉
         Case "501":
            'add by nickc 2006/06/30
            If textPrint = "1" Then
                ' 列印定稿
                'Modify By Cheng 2003/01/28
                '修改處理狀況
    '            NowPrint m_CP09, "06", "06", IIf(Me.textModLetter.Text = "Y", True, False), strUserNum, 0
'                NowPrint m_CP09, "06", "04", IIf(Me.textModLetter.Text = "Y", True, False), strUserNum, 0
               ET03 = "04" 'Modify By Sindy 2012/1/13
            End If
         ' 繳年費
         Case "708":
            'add by nickc 2006/06/30
            If textPrint = "1" Then
                ' 列印定稿
'                NowPrint m_CP09, "06", "07", IIf(Me.textModLetter.Text = "Y", True, False), strUserNum, 0
               ET03 = "07" 'Modify By Sindy 2012/1/13
            End If
      End Select
   End If
   
   'Add By Sindy 2012/1/13
   If ET03 <> "" Then
      bolEmail = PUB_GetEMailFlag(m_SP01 & m_SP02 & m_SP03 & m_SP04, , , bolPlusPaper)
      If bolEmail Then
         '判斷是否EMail同時寄紙本
         If Not bolPlusPaper Then
            iCopy = 1
         End If
         'Add By Sindy 2020/1/7 + 信函總收文號
         If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
            NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , , , , , , , strLD18
         Else
         '2020/1/7 END
            NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , True, True
            MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_SP01) & " ]！"
         End If
      Else
         'Add By Sindy 2019/12/20 + strLD18.信函總收文號
         NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , , , , , , , , , strLD18
      End If
   'Add By Sindy 2021/1/5 沒有系統產出的定稿
   Else
      'Add By Sindy 2021/2/1 詢問有沒有客戶函
      If strLD18 <> "" Then
         Call PUB_TCaseAskIsPost_C(strLD18)
      End If
   '2021/1/5 EMD
   End If
   '2012/1/13 End
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textCP64.Enabled = True Then
   Cancel = False
   textCP64_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textModLetter.Enabled = True Then
   Cancel = False
   textModLetter_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPeriod.Enabled = True Then
   Cancel = False
   textPeriod_Validate Cancel
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

If Me.textSP20.Enabled = True Then
   Cancel = False
   textSP20_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textSP21.Enabled = True Then
   Cancel = False
   textSP21_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function

