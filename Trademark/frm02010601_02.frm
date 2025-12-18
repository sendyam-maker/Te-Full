VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010601_02 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人已收達/已提申"
   ClientHeight    =   5748
   ClientLeft      =   48
   ClientTop       =   960
   ClientWidth     =   9348
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9348
   Begin VB.TextBox textResult 
      Height          =   264
      Left            =   780
      MaxLength       =   1
      TabIndex        =   0
      Text            =   "1"
      Top             =   5400
      Width           =   252
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1500
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   840
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   540
      Width           =   2532
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   840
      Width           =   2772
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6960
      TabIndex        =   2
      Top             =   72
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   435
      Left            =   5940
      TabIndex        =   1
      Top             =   60
      Width           =   972
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8220
      TabIndex        =   3
      Top             =   60
      Width           =   972
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   3552
      Left            =   72
      TabIndex        =   17
      Top             =   1752
      Width           =   9132
      _ExtentX        =   16108
      _ExtentY        =   6265
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
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   1500
      TabIndex        =   16
      Top             =   1140
      Width           =   7725
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13626;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1500
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1440
      Width           =   7725
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "13626;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
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
      Left            =   4110
      TabIndex        =   14
      Top             =   600
      Width           =   645
   End
   Begin VB.Label Label7 
      Caption         =   "(1:已收達 2:已提申)"
      Height          =   252
      Left            =   1140
      TabIndex        =   13
      Top             =   5400
      Width           =   1812
   End
   Begin VB.Label Label2 
      Caption         =   "結果 :"
      Height          =   252
      Left            =   180
      TabIndex        =   12
      Top             =   5400
      Width           =   492
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   180
      TabIndex        =   11
      Top             =   1440
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   252
      Left            =   180
      TabIndex        =   10
      Top             =   1140
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "申請案號 :"
      Height          =   252
      Index           =   13
      Left            =   180
      TabIndex        =   9
      Top             =   840
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   180
      TabIndex        =   8
      Top             =   540
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "審定號 :"
      Height          =   252
      Index           =   1
      Left            =   4800
      TabIndex        =   7
      Top             =   840
      Width           =   732
   End
End
Attribute VB_Name = "frm02010601_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/16 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Lydia 2021/08/17 改成Form2.0 ;textTM23、cmbTM05、grdList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/9 日期欄已修改
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
' 申請國家
Dim m_TM10 As String
Dim m_CurrSel As Integer
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

Private Sub cmdCancel_Click()
   Unload Me
   frm02010601_01.Show
End Sub

Private Sub cmdExit_Click()
   Unload frm02010601_01
   Unload Me
End Sub

Private Sub cmdok_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If IsEmptyText(m_CP09) = False Then
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
   textTM12.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
    'Modify By Cheng 2002/11/20
    '改在設計畫面預設
'   textResult = "1"
   
   MoveFormToCenter Me
   
   'Add By Sindy 2019/5/22
   m_strIR01 = frm02010601_01.m_strIR01
   m_strIR02 = frm02010601_01.m_strIR02
   m_strIR03 = frm02010601_01.m_strIR03
   m_strIR04 = frm02010601_01.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2019/5/22 END
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP05 = Empty
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
   End Select
End Sub

' 讀取商標基本檔
Private Sub QueryTradeMark()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   m_TM10 = Empty
   
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
      textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
      End If
      ' 審定號
      If IsNull(rsTmp.Fields("TM15")) = False Then
         textTM15 = rsTmp.Fields("TM15")
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
      ' 顯示案件名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
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
      textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
      ' 申請案號
      If IsNull(rsTmp.Fields("SP11")) = False Then
         textTM12 = rsTmp.Fields("SP11")
      End If
      ' 審定號
      If IsNull(rsTmp.Fields("SP14")) = False Then
         textTM15 = rsTmp.Fields("SP14")
      End If
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
      ' 顯示案件名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("SP08")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("SP08"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_TM10 = rsTmp.Fields("SP09")
      End If
      'add by nickc 2006/05/29 加入閉卷提示
      If IsNull(rsTmp.Fields("sp15")) Then
         Me.lblClose.Caption = ""
      Else
         Me.lblClose.Caption = "已閉卷"
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 查詢資料庫取得資料
Public Sub QueryData()
   
   Select Case m_TM01
      Case "T", "TF", "CFT", "FCT":
         ' 讀取商標基本檔
         QueryTradeMark
      Case Else:
         QueryServicePractice
   End Select
   
   ' 顯示符合條件的資料
   ListData
   
End Sub

' 列出案件進度表符合條件的資料
Private Sub ListData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim bDeal As Boolean
   Dim nIndex As Integer
   'Add By Cheng 2002/07/09
   Dim strTempName As String
   
   InitialGrdList
   
   m_CP09 = Empty
   
   'strSQL = "SELECT * FROM CaseProgress " & _
   '         "WHERE CP01 = '" & m_TM01 & "' AND " & _
   '               "CP02 = '" & m_TM02 & "' AND " & _
   '               "CP03 = '" & m_TM03 & "' AND " & _
   '               "CP04 = '" & m_TM04 & "' AND " & _
   '               "((CP46 = NULL AND CP47 = NULL) OR " & _
   '               "(CP46 <> NULL AND CP47 = NULL))"
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP01 = '" & m_TM01 & "' AND " & _
                  "CP02 = '" & m_TM02 & "' AND " & _
                  "CP03 = '" & m_TM03 & "' AND " & _
                  "CP04 = '" & m_TM04 & "' " & _
            "ORDER BY CP27 DESC "

   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         ' 無發文日不予列入
         If IsNull(rsTmp.Fields("CP27")) = True Then: GoTo NextRecord
         ' 發文日是空白不予列入
         If IsEmptyText(rsTmp.Fields("CP27")) = True Then: GoTo NextRecord
        'Modify By Cheng 2002/11/20
        '若結果為"1"已收達, 則收達日必須為Null
        If Me.textResult.Text = "1" Then
            ' 有代理人收達日或有代理人提申日不予計入
            If Not IsNull(rsTmp.Fields("CP46")) Then
               If Not IsEmptyText(rsTmp.Fields("CP46")) Then
                  GoTo NextRecord
               End If
            End If
        End If
        '若結果為"2"已提申, 則提申日必須為Null
        If Me.textResult.Text = "2" Then
            'Modify By Sindy 2010/12/23
            '結果為2時, 若操作人員為內商(部門P2X)則不限制提申日cp47條件
            If Left(GetStaffDepartment(strUserNum), 2) <> "P2" Then
            '2010/12/23 End
               ' 代理人提申日
               If Not IsNull(rsTmp.Fields("CP47")) Then
                  If Not IsEmptyText(rsTmp.Fields("CP47")) Then
                     GoTo NextRecord
                  End If
               End If
            End If
        End If
         ' 有實際結果的不予計入
         If Not IsNull(rsTmp.Fields("CP24")) Then
            If Not IsEmptyText(rsTmp.Fields("CP24")) Then
               GoTo NextRecord
            End If
         End If
         '92.4.24 CANCEL BY SONIA
         '' 有CF帳單編號的不予計入
         'If Not IsNull(rsTmp.Fields("CP61")) Then
         '   If Not IsEmptyText(rsTmp.Fields("CP61")) Then
         '      GoTo NextRecord
         '   End If
         'End If
         'If Not IsNull(rsTmp.Fields("CP62")) Then
         '   If Not IsEmptyText(rsTmp.Fields("CP62")) Then
         '      GoTo NextRecord
         '   End If
         'End If
         'If Not IsNull(rsTmp.Fields("CP63")) Then
         '   If Not IsEmptyText(rsTmp.Fields("CP63")) Then
         '      GoTo NextRecord
         '   End If
         'End If
         '92.4.24 END
         ' 無收文號不予列入
         If IsNull(rsTmp.Fields("CP09")) Then: GoTo NextRecord
         ' 收文號不為A,B類的不予計入
         Select Case Mid(rsTmp.Fields("CP09"), 1, 1)
            Case "A", "B":
            Case Else: GoTo NextRecord
         End Select
         'Add By Sindy 2018/5/14 踢除107.跨類 001.查名
         If m_TM01 = "CFT" Then
            If rsTmp.Fields("CP10") = "107" Or rsTmp.Fields("CP10") = "001" Then: GoTo NextRecord
         End If
         '2018/5/14 END
         
         ' 列入資料
         grdList.Rows = grdList.Rows + 1
         nIndex = grdList.Rows - 1
         ' 收文號
         If IsNull(rsTmp.Fields("CP09")) = False Then
            grdList.TextMatrix(nIndex, 1) = rsTmp.Fields("CP09")
         End If
         ' 收文日
         If IsNull(rsTmp.Fields("CP05")) = False Then
            If IsEmptyText(rsTmp.Fields("CP05")) = False And rsTmp.Fields("CP05") <> "0" Then
               grdList.TextMatrix(nIndex, 2) = TAIWANDATE(rsTmp.Fields("CP05"))
            End If
         End If
         ' 案件性質
         If IsNull(rsTmp.Fields("CP10")) = False Then
            If m_TM10 = "000" Then
               grdList.TextMatrix(nIndex, 3) = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
            Else
               grdList.TextMatrix(nIndex, 3) = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
            End If
         End If
         ' 2008/4/15 ADD BY SONIA 有相關總收文號時抓其案件性質
         If IsNull(rsTmp.Fields("CP43")) = False Then
            grdList.TextMatrix(nIndex, 3) = grdList.TextMatrix(nIndex, 3) & PUB_GetRelateCasePropertyName(grdList.TextMatrix(nIndex, 1), "1")
         End If
         '2008/4/15 END
         ' 發文日
         If IsNull(rsTmp.Fields("CP27")) = False Then
            If IsEmptyText(rsTmp.Fields("CP27")) = False And rsTmp.Fields("CP27") <> "0" Then
               grdList.TextMatrix(nIndex, 4) = TAIWANDATE(rsTmp.Fields("CP27"))
            End If
         End If
         ' 代理人
         If IsNull(rsTmp.Fields("CP44")) = False Then
            'Modify By Cheng 2002/07/09
'            grdList.TextMatrix(nIndex, 5) = GetFAgentName(rsTmp.Fields("CP44"))
            If PUB_GetAgentName(m_TM01, rsTmp.Fields("CP44"), strTempName) Then
               grdList.TextMatrix(nIndex, 5) = strTempName
            End If
         End If
NextRecord:
         rsTmp.MoveNext
      Loop
      'Added by Lydia 2023/10/16
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
      End If
      'end 2023/10/16
   End If
   
   ' 設定第一列為所選取的記錄
   grdList_SetSelection 1
   
End Sub

' 設定Grid List的一列為選取的狀態
Private Sub grdList_SetSelection(ByVal nSel As Integer)
   If grdList.Rows >= 2 Then
      grdList.row = nSel
      grdList_SelChange
      grdList_ShowSelection
   End If
End Sub

' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 6
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
   grdList.Text = "代理人"
   grdList.ColWidth(5) = 1200
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2019/5/22
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   'Add By Cheng 2002/07/18
   Set frm02010601_02 = Nothing
End Sub

Private Sub grdList_SelChange()
   If grdList.row > 0 Then
      m_CP09 = grdList.TextMatrix(grdList.row, 1)
   End If
   grdList_ShowSelection
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = grdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = grdList.row Then
      Dim nOldCol As Integer
      nOldCol = grdList.col
      grdList.col = 1
      If grdList.CellBackColor <> &H8000000D Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H8000000D Then grdList.CellBackColor = &H8000000D
            If grdList.CellForeColor <> &H80000005 Then grdList.CellForeColor = &H80000005
         Next nCol
      End If
      grdList.col = nOldCol
      GoTo EXITSUB
   End If
   
   ' 將原先選取的列回復到正常的顏色
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      If grdList.CellBackColor <> &H80000005 Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
            If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
         Next nCol
      End If
      grdList.col = 0
   End If
   ' 設定成所選取的列
   m_CurrSel = nCurrSel
   ' 將所選取的列反白
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      For nCol = 1 To grdList.Cols - 1
         grdList.col = nCol
         grdList.CellBackColor = &H8000000D
         grdList.CellForeColor = &H80000005
      Next nCol
      grdList.col = 0
   End If
EXITSUB:
End Sub

Private Sub DisplayNextForm()
   frm02010601_03.SetData 0, m_TM01, True
   frm02010601_03.SetData 1, m_TM02, False
   frm02010601_03.SetData 2, m_TM03, False
   frm02010601_03.SetData 3, m_TM04, False
   frm02010601_03.SetData 4, m_CP09, False
   'Add By Sindy 2019/5/22
   If Not m_PrevForm Is Nothing Then
      Call frm02010601_03.SetParent(m_PrevForm)
   End If
   frm02010601_03.m_strIR01 = m_strIR01
   frm02010601_03.m_strIR02 = m_strIR02
   frm02010601_03.m_strIR03 = m_strIR03
   frm02010601_03.m_strIR04 = m_strIR04
   '2019/5/22 END
   Me.Hide
   frm02010601_03.Show
   frm02010601_03.QueryData
End Sub

Private Sub textResult_Change()
    'Add By Cheng 2002/11/20
    ListData
End Sub

' 結果
Private Sub textResult_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If IsEmptyText(textResult) = False Then
      Select Case textResult
         Case "1", "2":
         Case Else
            strTit = "資料檢核"
            strMsg = "只可輸入1或2"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textResult_GotFocus
      End Select
   End If
End Sub

Public Function GetTextResult() As String
   GetTextResult = textResult
End Function

Private Sub textResult_GotFocus()
   InverseTextBox textResult
End Sub

