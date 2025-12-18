VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010411_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "智慧局註冊費通知函輸入"
   ClientHeight    =   3768
   ClientLeft      =   156
   ClientTop       =   972
   ClientWidth     =   9324
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3768
   ScaleWidth      =   9324
   Begin VB.TextBox textTM09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   540
      Width           =   2532
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7248
      TabIndex        =   4
      Top             =   45
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6420
      TabIndex        =   3
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8472
      TabIndex        =   5
      Top             =   45
      Width           =   800
   End
   Begin VB.TextBox textPrint 
      Height          =   264
      Left            =   5700
      MaxLength       =   1
      TabIndex        =   1
      Top             =   2670
      Width           =   372
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1250
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1470
      Width           =   2532
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2070
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1770
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1250
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1770
      Width           =   2532
   End
   Begin VB.TextBox textTM08 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1470
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1250
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   540
      Width           =   2532
   End
   Begin VB.TextBox textCP07 
      Height          =   264
      Left            =   1250
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2670
      Width           =   2532
   End
   Begin MSForms.TextBox TextCP13 
      Height          =   264
      Left            =   1250
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2070
      Width           =   2532
      VariousPropertyBits=   679493663
      ForeColor       =   0
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   300
      Left            =   1250
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   840
      Width           =   7992
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "14097;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP64 
      Height          =   492
      Left            =   1250
      TabIndex        =   2
      Top             =   3060
      Width           =   7992
      VariousPropertyBits=   -1467989989
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "14097;868"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   264
      Left            =   1250
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1170
      Width           =   8000
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "14111;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Caption         =   "智權人員："
      Height          =   255
      Left            =   300
      TabIndex        =   29
      Top             =   2070
      Width           =   915
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   9120
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label1 
      Caption         =   "商品類別："
      Height          =   255
      Index           =   1
      Left            =   4740
      TabIndex        =   26
      Top             =   540
      Width           =   900
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "(N:不通知)"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   6240
      TabIndex        =   25
      Top             =   2670
      Width           =   825
   End
   Begin VB.Label Label22 
      Caption         =   "是否通知客戶："
      Height          =   255
      Left            =   4400
      TabIndex        =   24
      Top             =   2670
      Width           =   1275
   End
   Begin VB.Label Label21 
      Caption         =   "進度備註："
      Height          =   255
      Left            =   300
      TabIndex        =   23
      Top             =   3060
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "審定號："
      Height          =   255
      Left            =   475
      TabIndex        =   22
      Top             =   1470
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日："
      Height          =   255
      Index           =   10
      Left            =   4565
      TabIndex        =   21
      Top             =   2070
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "彼所案號："
      Height          =   255
      Index           =   9
      Left            =   4740
      TabIndex        =   20
      Top             =   1770
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "通知函性質："
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   19
      Top             =   1770
      Width           =   1100
   End
   Begin VB.Label Label1 
      Caption         =   "商標種類："
      Height          =   255
      Index           =   2
      Left            =   4740
      TabIndex        =   18
      Top             =   1470
      Width           =   900
   End
   Begin VB.Label Label6 
      Caption         =   "申請人："
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   1170
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱："
      Height          =   255
      Left            =   300
      TabIndex        =   16
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   15
      Top             =   540
      Width           =   900
   End
   Begin VB.Label Label25 
      Caption         =   "最後期限："
      Height          =   255
      Left            =   300
      TabIndex        =   14
      Top             =   2670
      Width           =   900
   End
End
Attribute VB_Name = "frm02010411_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/29 Form2.0已修改 cbmTM05/textTM23/textCP13/textCP64
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/6 日期欄已修改
'2007/9/6 ADD BY SONIA  (因櫃台不輸此來函故不檢查來函記錄檔)
Option Explicit

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
Dim m_TM10 As String 'Add By Sindy 2009/09/24
' 來函收文日
Dim m_CP05 As String
' 是否已收文
Dim m_NP06 As String
' 通知函性質
Dim m_NP07 As String
' 法定期限
Dim m_NP09 As String
' 商標種類代碼
Dim m_TM08 As String
' 申請人國籍
Dim m_CU10 As String
' 原申請人代號
Dim m_TM23 As String
'Added by Morgan 2017/4/24 電子公文
Public m_DocWord As String
Public m_DocNo As String
Public m_AppNo As String
Public m_DeadLine As String
'end 2017/4/24
Dim m_TM30 As String 'Add By Sindy 2021/2/24 閉卷日期


Private Sub cmdCancel_Click()
   Unload Me
   If frm02010411_2.textTM12 = "" And frm02010411_2.textTM15 = "" Then
      Unload frm02010411_2
      frm02010411_1.Show
   Else
      frm02010411_2.Show
   End If
End Sub

Private Sub cmdExit_Click()
   Me.Enabled = False
   Unload frm02010411_1
   Unload frm02010411_2
   Unload Me
End Sub

Private Sub cmdok_Click()
   If CheckDataValid = True Then
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 儲存資料
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      Unload frm02010411_2
      'Added by Morgan 2017/4/24 電子公文
      If m_DocNo <> "" Then
         Unload Me
         Unload frm02010411_1
         frm02010412.GoNext
      Else
      'end 2017/4/24
         frm02010411_1.Show
         '2012/12/19 MODIFY BY SONIA 通知繳納第一期註冊費1715改為1720通知繳納註冊費
         If m_NP07 = "1720" Then
            frm02010411_1.textTM12.SetFocus
            frm02010411_1.textTM12 = ""
         Else
            frm02010411_1.textTM15.SetFocus
            frm02010411_1.textTM15 = ""
         End If
         Unload Me
      End If 'Added by Morgan 2017/4/24 電子公文
   End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM08.BackColor = &H8000000F
   textTM09.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   textTM45.BackColor = &H8000000F
   textCP05S.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
  
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
      m_NP06 = Empty
      m_NP07 = Empty
      m_NP09 = Empty
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
      ' 通知函性質
      Case 5: m_NP06 = strData
      ' 通知函性質
      Case 6: m_NP07 = strData
      ' 法定期限
      Case 7: m_NP09 = strData
   End Select
End Sub

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
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      '2012/12/19 MODIFY BY SONIA 通知繳納第一期註冊費1715改為1720通知繳納註冊費
      If m_NP07 = "1720" Then
         ' 申請案號
         Label2 = "申請案號 :"
         If IsNull(rsTmp.Fields("TM12")) = False Then
            textTM15 = rsTmp.Fields("TM12")
         End If
      Else
         ' 審定號
         Label2 = "審定號 :"
         If IsNull(rsTmp.Fields("TM15")) = False Then
            textTM15 = rsTmp.Fields("TM15")
         End If
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
         textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
      End If
      '商品類別
      Me.textTM09.Text = "" & rsTmp.Fields("TM09").Value
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         m_TM23 = rsTmp.Fields("TM23")
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
      ' 彼所案號
      If IsNull(rsTmp.Fields("TM45")) = False Then
         textTM45 = rsTmp.Fields("TM45")
      End If
      'Add By Sindy 2009/09/24
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
      End If
      'Add By Sindy 2021/2/24
      ' 閉卷日期
      m_TM30 = ""
      If IsNull(rsTmp.Fields("TM30")) = False Then
         m_TM30 = rsTmp.Fields("TM30")
      End If
      '2021/2/24 END
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Public Sub QueryData()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim strDay As String
   
   m_TM08 = Empty
   m_TM23 = Empty
   textPrint = Empty
   Label23.ForeColor = &H0&
   textPrint.Enabled = True
   
   ' 本所案號
   textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
   
   ' 來函收文日
   textCP05S = m_CP05
   
   ' 讀取商標基本檔的資料
   QueryTradeMark
      
   ' 通知函性質
   textCP10 = GetCaseTypeName(m_TM01, m_NP07, 0)
   ' 智權人員
   textCP13 = GetStaffName(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))
   'add by sonia 2015/10/12
   If textCP13 = "葉雪貞" Then
      MsgBox "此案為葉雪貞收文, 請注意！", vbExclamation
   'Modify by Amy 2017/03/07
   ElseIf Left(textCP13, 4) = "MCTF" Then
        MsgBox "此案為MCTF收文, 請注意！", vbExclamation
   'add by sonia 2018/7/23
   ElseIf textCP13 = "巨京商標" Then
        MsgBox "此案為巨京商標收文, 請注意！", vbExclamation
   End If
   'end 2015/10/12
   
   ' 已收文者不印通知函
   'Modify By Sindy 2021/2/24 +m_NP07="1717" 通知函性質=通知延展:MCT案已閉卷時,不通知客戶
   If m_NP06 = "Y" Or _
      (m_NP07 = "1717" And Left(textCP13, 4) = "MCTF" And Val(m_TM30) > 0) Then
      textPrint.Text = "N"
      textPrint.Enabled = False
      Label23.ForeColor = &HFF&
   End If
   
   'Added by Morgan 2017/4/24 電子公文
   If m_DeadLine <> "" Then
      If Len(m_DeadLine) >= 7 Then
         textCP07 = m_DeadLine
      End If
   End If
End Sub
   'end 2017/4/24

Public Function OnSaveData() As Boolean
Dim strSql As String
Dim strCP09 As String
Dim strCP27 As String
Dim strNP07 As String, strCP43 As String 'Add By Sindy 2020/1/31
Dim rsTmp As New ADODB.Recordset 'Add By Sindy 2020/1/31
   
On Error GoTo ErrorHandler
   OnSaveData = True
   cnnConnection.BeginTrans
      
   ' 新增資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   ' 已收文不必通知直接上發文日111111,要通知的不上發文日
   'If m_NP06 = "Y" Then
   'Modify By Sindy 2021/2/24 +m_NP07="1717" 通知函性質=通知延展:MCT案已閉卷時,不通知客戶
   If textPrint.Text = "N" Then
   '2021/2/24 END
      strCP27 = "19221111"
   Else
      strCP27 = ""
   End If
   
   '2007/9/14 add by sonia
   If textCP64 = "" Then
      textCP64 = "通知函最後繳費日:" & textCP07
   Else
      textCP64 = textCP64 & ",通知函最後繳費日:" & textCP07
   End If
   '2007/9/14 end
   
   'Add By Sindy 2020/1/31 讀取下一程序的總收文號
   Select Case m_NP07 '通知函性質
      Case "1720"
         strNP07 = "717"
      Case "1717"
         strNP07 = "102"
   End Select
   If strNP07 <> "" Then
      strSql = "SELECT NP01 FROM nextprogress " & _
               "WHERE NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND " & _
                     "NP07 = '" & strNP07 & "' AND " & _
                     "NP06 is null"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         strCP43 = rsTmp.Fields("NP01")
      End If
   End If
   '2020/1/31 END
   
   'Modify By Sindy 2020/1/31 + ,CP43 相關總收文號
   strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP64,CP43) " & _
            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                    "'" & strCP09 & "','" & m_NP07 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "'," & _
                    "'N','N','" & ChgSQL(strCP27) & "','N','" & ChgSQL(textCP64) & "'," & CNULL(strCP43) & ") "
   cnnConnection.Execute strSql
    
'   'Add By Sindy 2009/09/24
'   '因為有些來函由內商輸入，內商有自行控管之承辦期限及發文日。改為內商輸入所有C類來函，
'   '若業務區為F字頭者，除爭議受理外，自動產生B類收文，案件性質為外商發文722，不上發文日，不向客戶請款
'   Dim strCP48 As String, strCP09B As String
'   If Left(GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)), 1) = "F" And _
'      ((m_TM01 = "T" And m_TM10 = "020") Or (m_TM01 = "FCT" And m_TM10 = "000")) Then
'      strCP09B = AutoNo("B", 6)
'      '承辦期限為系統日加4個工作天
'      strCP48 = DBDATE(Pub_GetHandleDay(m_TM01, m_TM10, "722", strSrvDate(1), , m_CP09))
'      strSQL = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13,cp14,cp48,cp20,cp26,cp32,cp43) " & _
'                     "values (" & CNULL(m_TM01) & "," & CNULL(m_TM02) & "," & CNULL(m_TM03) & _
'                     "," & CNULL(m_TM04) & "," & CNULL(strSrvDate(1)) & "," & CNULL(strCP09B) & ",722," & _
'                     CNULL(m_CP12) & "," & CNULL(m_CP13) & "," & CNULL(m_CP13) & "," & CNULL(strCP48) & ",'N','N','N'," & CNULL(strCP09) & ")"
'      cnnConnection.Execute strSQL
'   End If
   
   'Added by Morgan 2017/4/24 電子公文
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, strCP09, m_TM01, m_TM02, m_TM03, m_TM04, m_NP07
   End If
   'end 2017/4/24
   
   cnnConnection.CommitTrans
   Exit Function

ErrorHandler:
    cnnConnection.RollbackTrans
    OnSaveData = False

End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm02010411_3 = Nothing
End Sub

' 最後期限
Private Sub textCP07_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP07) = False Then
      If CheckIsTaiwanDate(textCP07, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP07_GotFocus
         GoTo EXITSUB
      End If
   End If
EXITSUB:
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
      textCP64_GotFocus
   End If
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
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
   'Add by Amy 2021/12/29檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True) = False Then
        GoTo EXITSUB
   End If

   ' 法定期限不可空白
   If IsEmptyText(textCP07) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入最後期限"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP07.SetFocus
      GoTo EXITSUB
   Else
      ' 最後期限的日期不可小於系統日期
      If Val(Me.textCP07.Text) + 19110000 < ServerDate Then
         MsgBox "最後期限不可小於系統日期!!!", vbExclamation
         Me.textCP07.SetFocus
         textCP07_GotFocus
         GoTo EXITSUB
      End If
      '2012/12/19 MODIFY BY SONIA 通知繳納第一期註冊費1715改為1720通知繳納註冊費
      If m_NP07 = "1720" Then
         ' 第一期註冊費檢查必須大於或等於案件的法定期限
         '2012/9/20 modify by sonia 智慧局是以來函發文日計算期限,本所則以文到收文日次日計算,故有二天之差距
         'If Val(Me.textCP07.Text) + 19110000 < m_NP09 Then
         If Val(Me.textCP07.Text) + 19110000 < DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(m_NP09))) Then
            MsgBox "輸入的最後期限小於案件的法定期限!!! 請查明再輸 !", vbExclamation
            Me.textCP07.SetFocus
            textCP07_GotFocus
            GoTo EXITSUB
         End If
      Else
         ' 第二期註冊費及延展檢查必須與案件的法定期限相同
         If Val(Me.textCP07.Text) + 19110000 <> Val(m_NP09) Then
            MsgBox "輸入的最後期限與案件法定期限不同!!! 請查明再輸 !", vbExclamation
            Me.textCP07.SetFocus
            textCP07_GotFocus
            GoTo EXITSUB
         End If
      End If
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textCP07_GotFocus()
   InverseTextBox textCP07
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse

   TxtValidate = False
   If Me.textCP07.Enabled = True Then
      Cancel = False
      textCP07_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP64.Enabled = True Then
      Cancel = False
      textCP64_Validate Cancel
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

   TxtValidate = True
End Function
