VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010402_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "非爭議案核駁輸入"
   ClientHeight    =   5736
   ClientLeft      =   120
   ClientTop       =   996
   ClientWidth     =   9336
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   9336
   Begin VB.TextBox textTM06 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1620
      Width           =   7752
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1020
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Width           =   2532
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5520
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1020
      Width           =   2892
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7224
      TabIndex        =   2
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6396
      TabIndex        =   1
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8448
      TabIndex        =   3
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox textResult 
      Height          =   264
      Left            =   720
      MaxLength       =   1
      TabIndex        =   0
      Top             =   5400
      Width           =   252
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   2784
      Left            =   96
      TabIndex        =   20
      Top             =   2544
      Width           =   9072
      _ExtentX        =   16002
      _ExtentY        =   4911
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2220
      Width           =   7752
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "13674;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM07 
      Height          =   264
      Left            =   1440
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1920
      Width           =   7752
      VariousPropertyBits=   679493663
      Size            =   "13674;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM05 
      Height          =   264
      Left            =   1440
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1320
      Width           =   7752
      VariousPropertyBits=   679493663
      Size            =   "13674;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   19
      Top             =   2220
      Width           =   852
   End
   Begin VB.Label Label5 
      Caption         =   "案件日文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   1452
   End
   Begin VB.Label Label4 
      Caption         =   "案件英文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   17
      Top             =   1620
      Width           =   1212
   End
   Begin VB.Label Label3 
      Caption         =   "案件中文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "申請案號 :"
      Height          =   252
      Index           =   13
      Left            =   120
      TabIndex        =   15
      Top             =   1020
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "審定號 :"
      Height          =   252
      Index           =   1
      Left            =   4680
      TabIndex        =   13
      Top             =   1020
      Width           =   732
   End
   Begin VB.Label Label2 
      Caption         =   "結果 :"
      Height          =   252
      Left            =   120
      TabIndex        =   12
      Top             =   5400
      Width           =   492
   End
   Begin VB.Label Label7 
      Caption         =   "(1:核駁 2:改變原處份 3:申請駁回)"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   5400
      Width           =   3315
   End
End
Attribute VB_Name = "frm02010402_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/18 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Amy 2021/12/28 Form2.0已修改 textTM05/textTM07/textTM23/grdList
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
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
'Add By Sindy 2019/5/10
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2019/5/10 END


'Add By Sindy 2019/5/13
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdCancel_Click()
   Select Case strPrevForm
      Case "2"
         frm02010402_2.Show
         Unload Me
      Case Else
         frm02010402_1.Show
         Unload Me
         Unload frm02010402_2
   End Select
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
    '列印接洽接案單
'move to unload by nick 2004/10/22
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm02010402_2
   Unload frm02010402_1
   Unload Me
End Sub

Private Sub cmdok_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If IsEmptyText(m_CP09) = False Then
      If CheckDataValid() = True Then
         DisplayNextForm
      End If
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
   textTM12.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   ' 初始化
   Initial
   
   MoveFormToCenter Me
   
   'Add By Sindy 2019/5/10
   m_strIR01 = frm02010402_1.m_strIR01
   m_strIR02 = frm02010402_1.m_strIR02
   m_strIR03 = frm02010402_1.m_strIR03
   m_strIR04 = frm02010402_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2019/5/10 END
End Sub

Private Sub Initial()
   textResult = "1"
   'Added by Morgan 2017/4/17 電子公文
   Select Case frm02010402_1.m_NewCP10
      Case "1002"
         textResult = "1"
      Case "1403"
         textResult = "2"
   End Select
   'end 2017/4/17
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
                  " order by decode(cp10,'703',0,'704',0,cp27) desc,decode(cp10,'703',cp27,'704',cp27,99999999) desc "
            '"ORDER BY CP05 DESC "
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
         ' 收文號不為A,B類的不予計入
         Select Case Mid(rsTmp.Fields("CP09"), 1, 1)
            Case "A", "B":
            Case Else: GoTo NextRecord
         End Select
         ' 結果欄位 (結果欄位為 2 時表列出所有的資料)
         Select Case textResult
            ' 結果欄位為 1 時表列出無結果的資料
            Case "1":
               If IsNull(rsTmp.Fields("CP24")) = False Then
                  If IsEmptyText(rsTmp.Fields("CP24")) = False Then: GoTo NextRecord
               End If
            ' 結果欄位為 2 時表列出有結果的資料 modify by sonia 90.11.20
            Case "2":
               If IsNull(rsTmp.Fields("CP24")) = True Then: GoTo NextRecord
               If IsEmptyText(rsTmp.Fields("CP24")) = True Then: GoTo NextRecord
         End Select
         
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
            'modify by sonia 2017/7/19
            'grdList.TextMatrix(grdList.row, 3) = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
            If m_TM10 < "010" Then
               grdList.TextMatrix(grdList.row, 3) = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
            Else
               grdList.TextMatrix(grdList.row, 3) = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
            End If
            'end 2017/7/19
            ' 案件性質為救濟程序時需以紅色表示
            If Mid(rsTmp.Fields("CP10"), 1, 1) = "4" Then
               bSpecial = True
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
         'add by nickc 2008/01/22 後金
         If IsNull(rsTmp.Fields("CP19")) = False Then
            grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP19")
            If IsEmptyText(rsTmp.Fields("CP19")) = False Then
               bSpecial = True
            End If
         End If
'edit by nickc 2008/01/22 中間加入後金，往後移
'         ' 相關人
'         bDeal = False
'         If bDeal = False And IsNull(rsTmp.Fields("CP40")) = False Then
'            If IsEmptyText(rsTmp.Fields("CP40")) = False Then
'               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP40")
'               bDeal = True
'            End If
'         End If
'         If bDeal = False And IsNull(rsTmp.Fields("CP41")) = False Then
'            If IsEmptyText(rsTmp.Fields("CP41")) = False Then
'               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP41")
'               bDeal = True
'            End If
'         End If
'         If bDeal = False And IsNull(rsTmp.Fields("CP42")) = False Then
'            If IsEmptyText(rsTmp.Fields("CP42")) = False Then
'               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP42")
'               bDeal = True
'            End If
'         End If
'         If bDeal = False And IsNull(rsTmp.Fields("CP50")) = False Then
'            If IsEmptyText(rsTmp.Fields("CP50")) = False Then
'               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP50")
'               bDeal = True
'            End If
'         End If
'         If bDeal = False And IsNull(rsTmp.Fields("CP51")) = False Then
'            If IsEmptyText(rsTmp.Fields("CP51")) = False Then
'               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP51")
'               bDeal = True
'            End If
'         End If
'         If bDeal = False And IsNull(rsTmp.Fields("CP52")) = False Then
'            If IsEmptyText(rsTmp.Fields("CP52")) = False Then
'               grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP52")
'               bDeal = True
'            End If
'         End If
'         If bDeal = False And IsNull(rsTmp.Fields("CP56")) = False Then
'            If IsEmptyText(rsTmp.Fields("CP56")) = False Then
'               grdList.TextMatrix(grdList.row, 6) = GetCustomerName(rsTmp.Fields("CP56"), 0)
'               bDeal = True
'            End If
'         End If
'         If bDeal = False Then: grdList.TextMatrix(grdList.row, 6) = Empty
'         ' 特殊欄位
'         If bSpecial = True Then
'            grdList.TextMatrix(grdList.row, 7) = "1"
'         End If
         ' 相關人
         bDeal = False
         If bDeal = False And IsNull(rsTmp.Fields("CP40")) = False Then
            If IsEmptyText(rsTmp.Fields("CP40")) = False Then
               grdList.TextMatrix(grdList.row, 7) = rsTmp.Fields("CP40")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP41")) = False Then
            If IsEmptyText(rsTmp.Fields("CP41")) = False Then
               grdList.TextMatrix(grdList.row, 7) = rsTmp.Fields("CP41")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP42")) = False Then
            If IsEmptyText(rsTmp.Fields("CP42")) = False Then
               grdList.TextMatrix(grdList.row, 7) = rsTmp.Fields("CP42")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP50")) = False Then
            If IsEmptyText(rsTmp.Fields("CP50")) = False Then
               grdList.TextMatrix(grdList.row, 7) = rsTmp.Fields("CP50")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP51")) = False Then
            If IsEmptyText(rsTmp.Fields("CP51")) = False Then
               grdList.TextMatrix(grdList.row, 7) = rsTmp.Fields("CP51")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP52")) = False Then
            If IsEmptyText(rsTmp.Fields("CP52")) = False Then
               grdList.TextMatrix(grdList.row, 7) = rsTmp.Fields("CP52")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP56")) = False Then
            If IsEmptyText(rsTmp.Fields("CP56")) = False Then
               grdList.TextMatrix(grdList.row, 7) = GetCustomerName(rsTmp.Fields("CP56"), 0)
               bDeal = True
            End If
         End If
         If bDeal = False Then: grdList.TextMatrix(grdList.row, 7) = Empty
         ' 特殊欄位
         If bSpecial = True Then
            grdList.TextMatrix(grdList.row, 8) = "1"
         End If
NextRecord:
         rsTmp.MoveNext
      Loop
      'Added by Lydia 2023/10/18
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
      End If
      'end 2023/10/18
   End If
   
   'add by nickc 2008/01/22 加入 有後金時，顯示紅色
   Dim m_i As Integer
   Dim m_j As Integer
   For m_i = 1 To grdList.Rows - 1
        If Val(grdList.TextMatrix(m_i, 6)) > 0 Then
            grdList.row = m_i
            For m_j = 1 To grdList.Cols - 1
               If m_j = 6 Then
                   grdList.col = 6
                   grdList.CellBackColor = vbRed
               Else
                   grdList.col = m_j
                   grdList.CellBackColor = &H8000000D
                   grdList.CellForeColor = &H80000005
               End If
            Next m_j
        End If
   Next m_i
   
   ' 設定第一列為所選取的記錄
   grdList_SetSelection 1

End Sub

' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   'edit by nickc 2008/01/22 加入後金
   'grdList.Cols = 8
   grdList.Cols = 9
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
'edit by nickc 2008/01/22 加入後金
'   grdList.col = 6
'   grdList.Text = "相關人"
'   grdList.ColWidth(7) = 1200
'   grdList.col = 7
'   grdList.Text = "特殊"
'   grdList.ColWidth(7) = 0
   grdList.col = 6
   grdList.Text = "後金"
   grdList.ColWidth(6) = 1200
   grdList.col = 7
   grdList.Text = "相關人"
   grdList.ColWidth(7) = 1200
   grdList.col = 8
   grdList.Text = "特殊"
   grdList.ColWidth(8) = 0
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
            'add by nickc 2008/01/22 跳過6
            If nCol <> 6 Then
                grdList.col = nCol
                grdList.CellBackColor = &H8000000D
                grdList.CellForeColor = &H80000005
            End If
         Next nCol
      Else
         'If grdList.Cols = 7 Then
         '   grdList.col = 7
         If grdList.Cols = 9 Then
            grdList.col = 8
            Select Case grdList.Text
               Case 1:
                  grdList.col = 1
                  If grdList.CellBackColor <> &H80000005 Then
                     For nCol = 1 To grdList.Cols - 1
                        'add by nickc 2008/01/22 跳過6
                        If nCol <> 6 Then
                            grdList.col = nCol
                            If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
                            If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
                        End If
                     Next nCol
                  End If
               Case Else:
                  grdList.col = 1
                  If grdList.CellBackColor <> &H80000005 Then
                     For nCol = 1 To grdList.Cols - 1
                        'add by nickc 2008/01/22 跳過6
                        If nCol <> 6 Then
                            grdList.col = nCol
                            If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
                            If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
                        End If
                     Next nCol
                  End If
            End Select
         Else
            grdList.col = 1
            If grdList.CellBackColor <> &H80000005 Then
               For nCol = 1 To grdList.Cols - 1
                    'add by nickc 2008/01/22 跳過6
                    If nCol <> 6 Then
                        grdList.col = nCol
                        If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
                        If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
                    End If
               Next nCol
            End If
         End If
      End If
   Next nRow
   grdList.row = nCurrSel
   grdList.col = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
'edit by nickc 2008/04/25 改整批印
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   
   'Add By Sindy 2019/5/13
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   
   'Add By Cheng 2002/07/18
   Set frm02010402_3 = Nothing
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

Private Sub textResult_GotFocus()
   InverseTextBox textResult
End Sub

Private Sub textResult_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textResult) = False Then
      Select Case textResult
         'Modify By Cheng 2002/05/08
'         Case "1", "2":
         Case "1", "2", "3":
         Case Else
            Cancel = True
            strTit = "資料檢核"
'            strMsg = "請輸入 1 或 2 "
            strMsg = "請輸入 1 , 2 或 3 "
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textResult_GotFocus
            GoTo EXITSUB
      End Select
   Else
      textResult = "1"
   End If
   ListData
EXITSUB:
End Sub

Private Sub DisplayNextForm()
   frm02010402_4.SetData 0, m_TM01, True
   frm02010402_4.SetData 1, m_TM02, False
   frm02010402_4.SetData 2, m_TM03, False
   frm02010402_4.SetData 3, m_TM04, False
   frm02010402_4.SetData 4, m_CP05, False
   frm02010402_4.SetData 5, m_CP09, False
   'Added by Morgan 2017/4/12 電子公文
   frm02010402_4.m_DocWord = frm02010402_1.m_DocWord
   frm02010402_4.m_DocNo = frm02010402_1.m_DocNo
   frm02010402_4.m_AppNo = frm02010402_1.m_AppNo
   frm02010402_4.m_DeadLine = frm02010402_1.m_DeadLine
   'end 2017/4/12
   'Add By Sindy 2019/5/10
   If Not m_PrevForm Is Nothing Then
      Call frm02010402_4.SetParent(m_PrevForm)
   End If
   frm02010402_4.m_strIR01 = m_strIR01
   frm02010402_4.m_strIR02 = m_strIR02
   frm02010402_4.m_strIR03 = m_strIR03
   frm02010402_4.m_strIR04 = m_strIR04
   '2019/5/10 END
   frm02010402_4.PreResult = textResult 'Add by Amy 2022/09/26
   Me.Hide
   frm02010402_4.Show
   frm02010402_4.QueryData
End Sub

Public Function GetSelectResult() As String
   GetSelectResult = textResult
End Function

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   If IsEmptyText(textResult) Then
      strTit = "資料檢核"
      strMsg = "請輸入結果"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   Select Case textResult
      'Modify By Cheng 2002/05/08
'      Case "1", "2":
      Case "1", "2", "3":
      Case Else
         strTit = "資料檢核"
         strMsg = "結果欄位的值不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
   End Select
   
   CheckDataValid = True
EXITSUB:
End Function

