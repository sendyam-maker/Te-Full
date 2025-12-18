VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010503_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "被異議/被評定/被撤銷/對方補充理由/對方延期/通知復審答辯"
   ClientHeight    =   5688
   ClientLeft      =   240
   ClientTop       =   972
   ClientWidth     =   9156
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5688
   ScaleWidth      =   9156
   Begin VB.CommandButton cmdAllData 
      Caption         =   "來文資料(&A)"
      Height          =   400
      Left            =   4950
      TabIndex        =   22
      Top             =   70
      Width           =   1200
   End
   Begin VB.TextBox textResult 
      Height          =   264
      Left            =   720
      MaxLength       =   1
      TabIndex        =   0
      Top             =   5400
      Width           =   252
   End
   Begin VB.TextBox textTM06 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1560
      Width           =   7512
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   960
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   660
      Width           =   2532
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5640
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   660
      Width           =   2772
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7020
      TabIndex        =   2
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6192
      TabIndex        =   1
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8244
      TabIndex        =   3
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5640
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   2772
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   2772
      Left            =   48
      TabIndex        =   23
      Top             =   2472
      Width           =   8952
      _ExtentX        =   15790
      _ExtentY        =   4890
      _Version        =   393216
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
      _Band(0).Cols   =   2
   End
   Begin MSForms.TextBox textTM23 
      Height          =   264
      Left            =   1440
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2160
      Width           =   7512
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "13250;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM07 
      Height          =   264
      Left            =   1440
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1860
      Width           =   7512
      VariousPropertyBits=   679493663
      Size            =   "13250;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM05 
      Height          =   264
      Left            =   1440
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1260
      Width           =   7512
      VariousPropertyBits=   679493663
      Size            =   "13250;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   "結果 :"
      Height          =   252
      Left            =   120
      TabIndex        =   21
      Top             =   5400
      Width           =   492
   End
   Begin VB.Label Label7 
      Caption         =   "(1:被異議 2:被評定 3:被撤銷 4:對方補充理由 5:對方延期 6:通知復審答辨)"
      Height          =   252
      Left            =   1080
      TabIndex        =   20
      Top             =   5400
      Width           =   6492
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   19
      Top             =   2160
      Width           =   852
   End
   Begin VB.Label Label5 
      Caption         =   "案件日文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   18
      Top             =   1860
      Width           =   1452
   End
   Begin VB.Label Label4 
      Caption         =   "案件英文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   1212
   End
   Begin VB.Label Label3 
      Caption         =   "案件中文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   16
      Top             =   1260
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "申請案號 :"
      Height          =   252
      Index           =   13
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   660
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "審定號 :"
      Height          =   252
      Index           =   1
      Left            =   4680
      TabIndex        =   13
      Top             =   660
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家 :"
      Height          =   252
      Index           =   8
      Left            =   4680
      TabIndex        =   12
      Top             =   960
      Width           =   852
   End
End
Attribute VB_Name = "frm02010503_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/19 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Amy 2022/01/03 Form2.0已修改 textTM05/textTM07/textTM23/grdList
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/6 日期欄已修改
Option Explicit

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 來函收文日
Dim m_CP05 As String
' 所選取的收文號
Dim m_CP09 As String
' 來源畫面
Dim strPrevForm As String
' 申請國家
Dim m_TM10 As String
Dim m_TM14 As String  '2010/3/2 add by sonia
'Add By Sindy 2019/5/22
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2019/5/22 END


'Add By Sindy 2019/5/22
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdAllData_Click()
    'Add By Cheng 2002/12/17
    '若結果欄選擇對方補充理由或對方延期
    If Me.textResult.Text = "4" Or Me.textResult.Text = "5" Then
        ListData_1
    Else
        MsgBox "只有當結果欄為 4 或 5 時，來文資料按鈕才可使用!!!", vbExclamation + vbOKOnly
    End If
End Sub

Private Sub cmdCancel_Click()
   Select Case strPrevForm
      Case "2"
         frm02010503_2.Show
         Unload Me
      Case Else
         frm02010503_1.Show
         Unload frm02010503_2
         Unload Me
   End Select
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm02010503_2
   Unload frm02010503_1
   Unload Me
End Sub

Private Sub cmdok_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If IsEmptyText(m_CP09) = False Then
      '2010/3/2 ADD BY SONIA 大陸檢查被異議時必須有公告日,否則無法掛被異議續展期限
      'If m_TM10 = "020" And m_TM14 = "" Then
      If m_TM10 = "020" And m_TM14 = "" And m_TM01 <> "TD" And m_TM01 <> "TM" Then
         strMsg = "大陸案尚無公告日,不可輸入被異議,請先輸入核准資料!"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Exit Sub
      End If
      '2010/3/2 END
      DisplayNextForm
   Else
      strMsg = "請先選取一筆記錄"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM05.BackColor = &H8000000F
   textTM06.BackColor = &H8000000F
   textTM07.BackColor = &H8000000F
   textTM10.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   ' 初始化資料
   Initial
   
   MoveFormToCenter Me
   
   'Add By Sindy 2019/5/22
   m_strIR01 = frm02010503_1.m_strIR01
   m_strIR02 = frm02010503_1.m_strIR02
   m_strIR03 = frm02010503_1.m_strIR03
   m_strIR04 = frm02010503_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2019/5/22 END
End Sub

Private Sub Initial()
   textResult = "1"
   'Added by Morgan 2017/4/25 電子公文
   Select Case frm02010503_1.m_NewCP10
      Case "1601", "1602"
         textResult = "1"
      Case "1603", "1604"
         textResult = "2"
      Case "1605", "1606"
         textResult = "3"
      Case "1609"
         textResult = "4"
      Case "1611"
         textResult = "5"
      Case "1404"
         textResult = "6"
   End Select
   'end 2017/4/25
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP05 = Empty
      strPrevForm = Empty
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
      ' 來源畫面
      Case 5: strPrevForm = strData
   End Select
End Sub

' 查詢資料庫取得資料
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   m_TM10 = Empty: m_TM14 = ""
   
   ' 取得商標基本檔的相關項目
   '2011/6/9 ADD BY SONIA TM,TD自其他來函移過來
   Select Case m_TM01
      Case "TD", "TM":
         ' 設定SQL語法
         strSql = "SELECT SP01 AS TM01,SP02 AS TM02,SP03 AS TM03,SP04 AS TM04,SP05 AS TM05,SP06 AS TM06,SP07 AS TM07,SP09 AS TM10 " & _
            ",'' AS TM12,'' AS TM15,'' AS TM14,SP08 AS TM23 FROM SERVICEPRACTICE " & _
            "WHERE SP01 = '" & m_TM01 & "' AND " & _
                  "SP02 = '" & m_TM02 & "' AND " & _
                  "SP03 = '" & m_TM03 & "' AND " & _
                  "SP04 = '" & m_TM04 & "'"
      Case Else
   '2011/6/9 END
         strSql = "SELECT * FROM TradeMark " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "'"
   End Select  '2011/6/9 ADD BY SONIA
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
         textTM10 = GetNationName(rsTmp.Fields("TM10"), 0)
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
      End If
      ' 審定號
      If IsNull(rsTmp.Fields("TM15")) = False Then
         textTM15 = rsTmp.Fields("TM15")
      End If
      ' 商標名稱(中)
      If IsNull(rsTmp.Fields("TM05")) = False Then
         textTM05 = rsTmp.Fields("TM05")
      End If
      ' 商標名稱(英)
      If IsNull(rsTmp.Fields("TM06")) = False Then
         textTM06 = rsTmp.Fields("TM06")
      End If
      ' 商標名稱(日)
      If IsNull(rsTmp.Fields("TM07")) = False Then
         textTM07 = rsTmp.Fields("TM07")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
      ' 公告日 2010/3/2 add by sonia
      If IsNull(rsTmp.Fields("TM14")) = False Then
         m_TM14 = rsTmp.Fields("TM14")
      End If
   End If
   rsTmp.Close
   ' 顯示符合條件的資料
   ListData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 列出案件進度表符合條件的資料
Private Sub ListData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim bDeal As Boolean
   Dim bSpecial As Boolean
   
   InitialGrdList
   
   m_CP09 = Empty
   
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP01 = '" & m_TM01 & "' AND " & _
                  "CP02 = '" & m_TM02 & "' AND " & _
                  "CP03 = '" & m_TM03 & "' AND " & _
                  "CP04 = '" & m_TM04 & "' " & _
            "ORDER BY CP05 DESC "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         bSpecial = False
         ' 無發文日不予列入
         If IsNull(rsTmp.Fields("CP27")) = True Then: GoTo NextRecord
         ' 發文日是空白不予列入
         If IsEmptyText(rsTmp.Fields("CP27")) = True Then: GoTo NextRecord
         ' 無收文號不予列入
         If IsNull(rsTmp.Fields("CP09")) = True Then: GoTo NextRecord
         ' 收文號不為A,B,C類的不予計入
         Select Case Mid(rsTmp.Fields("CP09"), 1, 1)
'            Case "A", "B":
            Case "A", "B", "C":
            Case Else: GoTo NextRecord
         End Select
        'Add By Cheng 2004/02/20
'         ' C類只取 被異議(理由), 被評定(理由), 被廢止(理由), 通知參加訴願, 通知參加訴訟, 通知言詞辯論, 通知準備程序
         ' C類只取 被異議, 被異議(理由), 被評定, 被評定(理由), 被廢止, 被廢止(理由), 通知參加訴願, 通知參加訴訟, 通知言詞辯論, 通知準備程序, 通知行政上訴答辯
         If Mid(rsTmp.Fields("CP09"), 1, 1) = "C" Then
            If IsNull(rsTmp.Fields("CP10")) = False Then
               Select Case rsTmp.Fields("CP10")
'                  Case "1602", "1604", "1606", "1404", "1405", "1203", "1204":
                  Case "1601", "1602", "1603", "1604", "1605", "1606", "1404", "1405", "1203", "1204", "1406":
                  Case Else: GoTo NextRecord
               End Select
            End If
         End If
        'End
         ' 列入資料
         grdList.Rows = grdList.Rows + 1
         grdList.row = grdList.Rows - 1
         ' 收文號
         If IsNull(rsTmp.Fields("CP09")) = False Then
            grdList.TextMatrix(grdList.row, 1) = rsTmp.Fields("CP09")
         End If
         ' 收文日
         If IsNull(rsTmp.Fields("CP05")) = False Then
            grdList.TextMatrix(grdList.row, 2) = TAIWANDATE(rsTmp.Fields("CP05"))
         End If
         ' 案件性質
         If IsNull(rsTmp.Fields("CP10")) = False Then
            If m_TM10 = "000" Then
               grdList.TextMatrix(grdList.row, 3) = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
            Else
               grdList.TextMatrix(grdList.row, 3) = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
            End If
         End If
         ' 發文日
         If IsNull(rsTmp.Fields("CP27")) = False Then
            grdList.TextMatrix(grdList.row, 4) = TAIWANDATE(rsTmp.Fields("CP27"))
         End If
         ' 結果
         If IsNull(rsTmp.Fields("CP24")) = False Then
            Select Case rsTmp.Fields("CP24")
               Case "1":
                  grdList.TextMatrix(grdList.row, 5) = "准/勝"
               Case "2":
                  grdList.TextMatrix(grdList.row, 5) = "駁/敗"
               Case Else:
            End Select
         End If
         ' 相關人
         bDeal = False
         If bDeal = False And IsNull(rsTmp.Fields("CP40")) = False Then
            If IsEmptyText(rsTmp.Fields("CP40")) = False Then
               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP40")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP41")) = False Then
            If IsEmptyText(rsTmp.Fields("CP41")) = False Then
               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP41")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP42")) = False Then
            If IsEmptyText(rsTmp.Fields("CP42")) = False Then
               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP42")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP50")) = False Then
            If IsEmptyText(rsTmp.Fields("CP50")) = False Then
               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP50")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP51")) = False Then
            If IsEmptyText(rsTmp.Fields("CP51")) = False Then
               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP51")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP52")) = False Then
            If IsEmptyText(rsTmp.Fields("CP52")) = False Then
               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP52")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP56")) = False Then
            If IsEmptyText(rsTmp.Fields("CP56")) = False Then
               grdList.TextMatrix(grdList.row, 6) = GetCustomerName(rsTmp.Fields("CP56"), 0)
               bDeal = True
            End If
         End If
         If bDeal = False Then: grdList.TextMatrix(grdList.row, 6) = Empty
         ' 特殊欄位
         If bSpecial = True Then
            grdList.TextMatrix(grdList.row, 7) = "1"
         End If
NextRecord:
         rsTmp.MoveNext
      Loop
      
      'Added by Lydia 2023/10/19
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
      End If
      'end 2023/10/19
   End If
   
   ' 設定第一列為所選取的記錄
   grdList_SetSelection 1

End Sub

'Add By Cheng 2002/12/17
'For 對方補充理由或對方延期
' 列出案件進度表符合條件的資料
Private Sub ListData_1()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim bDeal As Boolean
   Dim bSpecial As Boolean
   
   InitialGrdList
   
   m_CP09 = Empty
   
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP01 = '" & m_TM01 & "' AND " & _
                  "CP02 = '" & m_TM02 & "' AND " & _
                  "CP03 = '" & m_TM03 & "' AND " & _
                  "CP04 = '" & m_TM04 & "' " & _
            "ORDER BY CP05 DESC "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         bSpecial = False
         ' 無收文號不予列入
         If IsNull(rsTmp.Fields("CP09")) = True Then: GoTo NextRecord
        'Add By Cheng 2004/02/20
         ' 收文號不為A,B,C類的不予計入
         Select Case Mid(rsTmp.Fields("CP09"), 1, 1)
            Case "A", "B", "C":
            Case Else: GoTo NextRecord
         End Select
'         ' C類只取 被異議(理由), 被評定(理由), 被廢止(理由), 通知參加訴願, 通知參加訴訟, 通知言詞辯論, 通知準備程序
         ' C類只取 被異議, 被異議(理由), 被評定, 被評定(理由), 被廢止, 被廢止(理由), 通知參加訴願, 通知參加訴訟, 通知言詞辯論, 通知準備程序
         If Mid(rsTmp.Fields("CP09"), 1, 1) = "C" Then
            If IsNull(rsTmp.Fields("CP10")) = False Then
               'Modify By Sindy 2024/9/26 改為常變數
               If InStr(商爭審查來函案件性質, rsTmp.Fields("CP10")) = 0 Then
                  GoTo NextRecord
               End If
'               Select Case rsTmp.Fields("CP10")
''                  Case "1602", "1604", "1606", "1404", "1405", "1203", "1204":
'                  Case "1601", "1602", "1603", "1604", "1605", "1606", "1404", "1405", "1203", "1204", "1406":
'                  '2007/5/29 ADD BY SONIA 增1205部分核駁,1609對方補充理由,1612發回補理由,1613發回補答辯,1616通知復審答辯,1618對方答辯,1619被部分廢止,1620被部分廢止(理由),1621對造分割
'                  Case "1205", "1609", "1612", "1613", "1616", "1618", "1619", "1620", "1621"
'                  Case Else: GoTo NextRecord
'               End Select
               '2024/9/26 END
            End If
         End If
        'End
         ' 列入資料
         grdList.Rows = grdList.Rows + 1
         grdList.row = grdList.Rows - 1
         ' 收文號
         If IsNull(rsTmp.Fields("CP09")) = False Then
            grdList.TextMatrix(grdList.row, 1) = rsTmp.Fields("CP09")
         End If
         ' 收文日
         If IsNull(rsTmp.Fields("CP05")) = False Then
            grdList.TextMatrix(grdList.row, 2) = TAIWANDATE(rsTmp.Fields("CP05"))
         End If
         ' 案件性質
         If IsNull(rsTmp.Fields("CP10")) = False Then
            If m_TM10 = "000" Then
               grdList.TextMatrix(grdList.row, 3) = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
            Else
               grdList.TextMatrix(grdList.row, 3) = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
            End If
         End If
         ' 發文日
         If IsNull(rsTmp.Fields("CP27")) = False Then
            grdList.TextMatrix(grdList.row, 4) = TAIWANDATE(rsTmp.Fields("CP27"))
         End If
         ' 結果
         If IsNull(rsTmp.Fields("CP24")) = False Then
            Select Case rsTmp.Fields("CP24")
               Case "1":
                  grdList.TextMatrix(grdList.row, 5) = "准/勝"
               Case "2":
                  grdList.TextMatrix(grdList.row, 5) = "駁/敗"
               Case Else:
            End Select
         End If
         ' 相關人
         bDeal = False
         If bDeal = False And IsNull(rsTmp.Fields("CP40")) = False Then
            If IsEmptyText(rsTmp.Fields("CP40")) = False Then
               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP40")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP41")) = False Then
            If IsEmptyText(rsTmp.Fields("CP41")) = False Then
               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP41")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP42")) = False Then
            If IsEmptyText(rsTmp.Fields("CP42")) = False Then
               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP42")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP50")) = False Then
            If IsEmptyText(rsTmp.Fields("CP50")) = False Then
               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP50")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP51")) = False Then
            If IsEmptyText(rsTmp.Fields("CP51")) = False Then
               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP51")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP52")) = False Then
            If IsEmptyText(rsTmp.Fields("CP52")) = False Then
               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP52")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP56")) = False Then
            If IsEmptyText(rsTmp.Fields("CP56")) = False Then
               grdList.TextMatrix(grdList.row, 6) = GetCustomerName(rsTmp.Fields("CP56"), 0)
               bDeal = True
            End If
         End If
         If bDeal = False Then: grdList.TextMatrix(grdList.row, 6) = Empty
         ' 特殊欄位
         If bSpecial = True Then
            grdList.TextMatrix(grdList.row, 7) = "1"
         End If
NextRecord:
         rsTmp.MoveNext
      Loop
      
      'Added by Lydia 2023/10/19
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
      End If
      'end 2023/10/19
   End If
   
   ' 設定第一列為所選取的記錄
   grdList_SetSelection 1

End Sub

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
   Set frm02010503_3 = Nothing
End Sub

Private Sub textResult_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textResult) = False Then
      Select Case textResult
         Case "1", "2", "3", "4", "5":
         Case "6":
            If m_TM10 < "010" Then
               Cancel = True
               strTit = "資料檢核"
               strMsg = "申請國家為台灣, 不可選擇通知復審答辯"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textResult_GotFocus
               GoTo EXITSUB
            End If
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "請輸入1 或 2 或 3 或 4 或 5 或 6"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textResult_GotFocus
            GoTo EXITSUB
      End Select
   Else
      textResult = "1"
   End If
   'ListData
EXITSUB:
End Sub

' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 8
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "收文號"
   grdList.ColWidth(1) = 1200
   grdList.col = 2
   grdList.Text = "收文日"
   grdList.ColWidth(2) = 1000
   grdList.col = 3
   grdList.Text = "案件性質"
   grdList.ColWidth(3) = 1200
   grdList.col = 4
   grdList.Text = "發文日"
   grdList.ColWidth(4) = 1000
   grdList.col = 5
   grdList.Text = "結果"
   grdList.ColWidth(5) = 1200
   grdList.col = 6
   grdList.Text = "相關人"
   grdList.ColWidth(6) = 1200
   grdList.col = 7
   grdList.Text = "特殊"
   grdList.ColWidth(7) = 0
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nRow As Integer
   Dim nCol As Integer
   Dim nCurrSel As Integer
   nCurrSel = grdList.row
   For nRow = 1 To grdList.Rows - 1
      grdList.row = nRow
      If nRow = nCurrSel Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            grdList.CellBackColor = &H8000000D
            grdList.CellForeColor = &H80000005
         Next nCol
      Else
         If grdList.Cols = 9 Then
            grdList.col = 8
            Select Case grdList.Text
               Case 1:
                  grdList.col = 1
                  If grdList.CellBackColor <> &HFF& Then
                     For nCol = 1 To grdList.Cols - 1
                        grdList.col = nCol
                        If grdList.CellBackColor <> &HFF& Then: grdList.CellBackColor = &HFF&
                        If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
                     Next nCol
                  End If
               Case Else:
                  grdList.col = 1
                  If grdList.CellBackColor <> &H80000005 Then
                     For nCol = 1 To grdList.Cols - 1
                        grdList.col = nCol
                        If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
                        If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
                     Next nCol
                  End If
            End Select
         Else
            grdList.col = 1
            If grdList.CellBackColor <> &H80000005 Then
               For nCol = 1 To grdList.Cols - 1
                  grdList.col = nCol
                  If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
                  If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
               Next nCol
            End If
         End If
      End If
   Next nRow
   grdList.row = nCurrSel
   grdList.col = 0
End Sub

Private Sub grdList_SelChange()
   If grdList.row > 0 Then
      grdList.col = 1
      m_CP09 = grdList.Text
   End If
   grdList_ShowSelection
End Sub

' 設定Grid List的一列為選取的狀態
Private Sub grdList_SetSelection(ByVal nSel As Integer)
   If grdList.Rows >= 2 Then
      grdList.row = nSel
      grdList_SelChange
      grdList_ShowSelection
   End If
End Sub

Private Sub DisplayNextForm()
   frm02010503_4.SetData 0, m_TM01, True
   frm02010503_4.SetData 1, m_TM02, False
   frm02010503_4.SetData 2, m_TM03, False
   frm02010503_4.SetData 3, m_TM04, False
   frm02010503_4.SetData 4, m_CP05, False
   frm02010503_4.SetData 5, m_CP09, False
   'Added by Morgan 2017/4/25 電子公文
   frm02010503_4.m_DocWord = frm02010503_1.m_DocWord
   frm02010503_4.m_DocNo = frm02010503_1.m_DocNo
   frm02010503_4.m_AppNo = frm02010503_1.m_AppNo
   frm02010503_4.m_DeadLine = frm02010503_1.m_DeadLine
   'end 2017/4/25
   'Add By Sindy 2019/5/22
   If Not m_PrevForm Is Nothing Then
      Call frm02010503_4.SetParent(m_PrevForm)
   End If
   frm02010503_4.m_strIR01 = m_strIR01
   frm02010503_4.m_strIR02 = m_strIR02
   frm02010503_4.m_strIR03 = m_strIR03
   frm02010503_4.m_strIR04 = m_strIR04
   '2019/5/22 END
   Me.Hide
   frm02010503_4.Show
   frm02010503_4.QueryData
End Sub

Public Function GetSelectResult() As String
   GetSelectResult = textResult
End Function

Private Sub textResult_GotFocus()
   InverseTextBox textResult
End Sub

