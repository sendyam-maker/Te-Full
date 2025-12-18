VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03020401_03 
   BorderStyle     =   1  '單線固定
   Caption         =   "非爭議案核准輸入"
   ClientHeight    =   5748
   ClientLeft      =   4176
   ClientTop       =   2388
   ClientWidth     =   9324
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9324
   Begin VB.TextBox textResult 
      Height          =   264
      Left            =   780
      MaxLength       =   1
      TabIndex        =   0
      Top             =   5400
      Width           =   252
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8280
      TabIndex        =   3
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6060
      TabIndex        =   1
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7020
      TabIndex        =   2
      Top             =   60
      Width           =   1212
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   900
      Width           =   2772
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Width           =   2532
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   900
      Width           =   2532
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   3555
      Left            =   90
      TabIndex        =   17
      Top             =   1770
      Width           =   9135
      _ExtentX        =   16108
      _ExtentY        =   6287
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
      Height          =   285
      Left            =   1440
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1500
      Width           =   7725
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "13626;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   1440
      TabIndex        =   15
      Top             =   1200
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
      Height          =   285
      Left            =   4080
      TabIndex        =   14
      Top             =   630
      Width           =   645
   End
   Begin VB.Label Label2 
      Caption         =   "結果 :"
      Height          =   252
      Left            =   180
      TabIndex        =   13
      Top             =   5400
      Width           =   492
   End
   Begin VB.Label Label7 
      Caption         =   "(1:核准 2:改變原處分)"
      Height          =   252
      Left            =   1140
      TabIndex        =   12
      Top             =   5400
      Width           =   2892
   End
   Begin VB.Label Label1 
      Caption         =   "審定號 :"
      Height          =   252
      Index           =   1
      Left            =   4920
      TabIndex        =   11
      Top             =   900
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "申請案號 :"
      Height          =   252
      Index           =   13
      Left            =   120
      TabIndex        =   9
      Top             =   900
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1332
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   7
      Top             =   1500
      Width           =   852
   End
End
Attribute VB_Name = "frm03020401_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/03/18 grdList : MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉
'Memo by Lydia 2021/09/13 改成Form2.0 ; cmbTM05、textTM23、grdList改字型=新細明體-ExtB
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
' 所選取的收文號
Dim m_CP09 As String
' 申請國家
Dim m_TM10 As String
'
Dim m_CurrSel As Integer
' 2006/6/1 ADD BY SONIA 所選取的案件性質
Dim m_CP10 As String
'add by nickc 2006/07/21  分割母案控制
Dim Is308Monther As Boolean
Dim IsHaveTM15 As Boolean


Private Sub cmdCancel_Click()
   Unload Me
   frm03020401_02.Show
End Sub

Private Sub cmdExit_Click()
   Unload frm03020401_02
   Unload frm03020401_01
   Unload Me
End Sub

Private Sub cmdOK_Click()
      Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If IsEmptyText(m_CP09) = False Then
      '2006/6/1 MODIFY BY SONIA 由frm03020401_02移至此檢查,因為申請有第一期註冊費期限,其他案件性質無期限
      'DisplayNextForm
      If PromptIfTaiwanNoResult = True Then
         'add by nickc 2006/07/21
         If ChkTmData = True Then
            DisplayNextForm
         End If
      Else
         Unload Me
         frm03020401_01.Show
      End If
      '2006/6/1 END
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
   
   ' 設定初始值
   Initial
   
   MoveFormToCenter Me
End Sub

Private Sub Initial()
   textResult = "1"
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
      textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
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

' 查詢資料庫取得資料
Public Sub QueryData()
   
   ' 讀取商標基本檔
   QueryTradeMark
   
   ' 顯示符合條件的資料
   ListData
End Sub

' 列出案件進度表符合條件的資料
Private Sub ListData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim bDeal As Boolean
   Dim nIndex As Integer
   
   InitialGrdList
   
   m_CP09 = Empty: m_CP10 = Empty
   
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP01 = '" & m_TM01 & "' AND " & _
                  "CP02 = '" & m_TM02 & "' AND " & _
                  "CP03 = '" & m_TM03 & "' AND " & _
                  "CP04 = '" & m_TM04 & "' "
'add by nick 2004/10/18 只有 101,102,103,301,302,304,306,307,308,309,310,501,502,503,504,505,506,507,716  才抓
'2007/8/9 modify by sonia 加313減縮商品
'2009/10/14 modify by sonia 加724徵求同意書
'2009/11/12 MODIFY BY SONIA 加725退費
'modify by sonia 2022/9/28 暫緩審理310排除，改在延期受理輸
'Modified by Morgan 2023/10/11 +314申請註冊證副本
   strSql = strSql & " and cp10 in ('101','102','103','301','302','304','306','307','308','309','314','501','502','503','504','505','506','507','716','313','724','725') "
   '2008/10/1 ADD BY SONIA 加排序條件
   strSql = strSql & " ORDER BY CP05,CP09 "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
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
         
         ' 結果欄位 (結果欄位為 2 時表列出所有的資料) copy by sonia 90.11.20
         Select Case textResult
            ' 結果欄位為 1 時表列出無結果的資料
            Case "1":
               If IsNull(rsTmp.Fields("CP24")) = False Then
                  If IsEmptyText(rsTmp.Fields("CP24")) = False Then: GoTo NextRecord
               End If
            ' 結果欄位為 2 時表列出有結果的資料
            Case "2":
               If IsNull(rsTmp.Fields("CP24")) = True Then: GoTo NextRecord
               If IsEmptyText(rsTmp.Fields("CP24")) = True Then: GoTo NextRecord
         End Select
         
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
               grdList.TextMatrix(nIndex, 2) = rsTmp.Fields("CP05")
            End If
         End If
         ' 案件性質
         If IsNull(rsTmp.Fields("CP10")) = False Then
            grdList.TextMatrix(nIndex, 3) = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         End If
         ' 發文日
         If IsNull(rsTmp.Fields("CP27")) = False Then
            If IsEmptyText(rsTmp.Fields("CP27")) = False And rsTmp.Fields("CP27") <> "0" Then
               grdList.TextMatrix(nIndex, 4) = rsTmp.Fields("CP27")
            End If
         End If
         ' 結果
         If IsNull(rsTmp.Fields("CP24")) = False Then
            Select Case rsTmp.Fields("CP24")
               Case "1":
                  grdList.TextMatrix(nIndex, 5) = "准/勝"
               Case "2":
                  grdList.TextMatrix(nIndex, 5) = "駁/敗"
               Case Else:
            End Select
         End If
         ' 相關人
         bDeal = False
         If bDeal = False And IsNull(rsTmp.Fields("CP40")) = False Then
            If IsEmptyText(rsTmp.Fields("CP40")) = False Then
               grdList.TextMatrix(nIndex, 6) = rsTmp.Fields("CP40")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP41")) = False Then
            If IsEmptyText(rsTmp.Fields("CP41")) = False Then
               grdList.TextMatrix(nIndex, 6) = rsTmp.Fields("CP41")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP42")) = False Then
            If IsEmptyText(rsTmp.Fields("CP42")) = False Then
               grdList.TextMatrix(nIndex, 6) = rsTmp.Fields("CP42")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP50")) = False Then
            If IsEmptyText(rsTmp.Fields("CP50")) = False Then
               grdList.TextMatrix(nIndex, 6) = rsTmp.Fields("CP50")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP51")) = False Then
            If IsEmptyText(rsTmp.Fields("CP51")) = False Then
               grdList.TextMatrix(nIndex, 6) = rsTmp.Fields("CP51")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP52")) = False Then
            If IsEmptyText(rsTmp.Fields("CP52")) = False Then
               grdList.TextMatrix(nIndex, 6) = rsTmp.Fields("CP52")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP56")) = False Then
            If IsEmptyText(rsTmp.Fields("CP56")) = False Then
               grdList.TextMatrix(nIndex, 6) = GetCustomerName(rsTmp.Fields("CP56"), 0)
               bDeal = True
            End If
         End If
         'Modifed by Lydia 2022/03/18 因為會直接清掉"相關人"的標題
         'If bDeal = False Then: grdList.Text = Empty
         If bDeal = False Then grdList.TextMatrix(nIndex, 6) = ""
NextRecord:
         rsTmp.MoveNext
      Loop
   End If
   
    'Added by Lydia 2022/03/18 MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉
    If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
    End If
    'end 2022/03/18
    
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
   grdList.Cols = 7
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/19
   Set frm03020401_03 = Nothing
End Sub

Private Sub grdList_SelChange()
   If grdList.row > 0 Then
      m_CP09 = grdList.TextMatrix(grdList.row, 1)
      '2006/6/1 ADD BY SONIA
      m_CP10 = grdList.TextMatrix(grdList.row, 3)
      '2006/6/1 END
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
    'add by nickc 2006/07/21
    If Is308Monther = False Then
        frm03020401_04.SetData 0, m_TM01, True
        frm03020401_04.SetData 1, m_TM02, False
        frm03020401_04.SetData 2, m_TM03, False
        frm03020401_04.SetData 3, m_TM04, False
        frm03020401_04.SetData 4, m_CP05, False
        frm03020401_04.SetData 5, m_CP09, False
      
         'Added by Morgan 2017/5/3 電子公文
         frm03020401_04.m_DocWord = frm03020401_01.m_DocWord
         frm03020401_04.m_DocNo = frm03020401_01.m_DocNo
         frm03020401_04.m_AppNo = frm03020401_01.m_AppNo
         frm03020401_04.m_DeadLine = frm03020401_01.m_DeadLine
         'end 2017/5/3
      
        Me.Hide
        frm03020401_04.Show
        
        'Add By Cheng 2002/02/01
        frm03020401_04.SetLastData
        
        frm03020401_04.QueryData
    'add by nickc 2006/07/21
    Else
        frm02010401_6.oKey = m_CP09
        frm02010401_6.IsHaveTM15 = IsHaveTM15
        frm02010401_6.oStrCDate = frm03020401_01.textCP05
        Set frm02010401_6.UpForm = Me
        
         'Added by Morgan 2017/5/3 電子公文
         frm02010401_6.m_DocWord = frm03020401_01.m_DocWord
         frm02010401_6.m_DocNo = frm03020401_01.m_DocNo
         frm02010401_6.m_AppNo = frm03020401_01.m_AppNo
         frm02010401_6.m_DeadLine = frm03020401_01.m_DeadLine
         'end 2017/5/3
         
        Me.Hide
        frm02010401_6.Show
        frm02010401_6.StrMenu
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
         Case "1", "2":
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "請輸入1 或 2"
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

Public Function GetSelectResult() As String
   GetSelectResult = textResult
End Function
'2006/6/1 ADD BY SONIA 由frm03020401_02移至此檢查,因為申請有第一期註冊費期限,其他案件性質無期限
' 檢查來函記錄檔
Private Function PromptIfTaiwanNoResult() As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim bPrompt As Boolean
   Dim strDate As String

   bPrompt = False
   PromptIfTaiwanNoResult = True
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("TM10")) = False Then
         If rsTmp.Fields("TM10") < "010" Then
            'modify by sonia 2013/7/16 分割不一定有無期限
            'If m_CP10 <> "申請" Then   '非申請案無期限
            If m_CP10 = "分割" Then
            ElseIf m_CP10 <> "申請" Then   '非申請案無期限
            '2013/7/16 end
               If IsMailRecNoTermExist(m_TM01, m_TM02, m_TM03, m_TM04, m_CP05) = False Then
                  bPrompt = True
               End If
            Else
               strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, m_CP05, "MR16")
               If IsEmptyText(strDate) = True Then
                  bPrompt = True
               End If
            End If
         End If
      End If
   End If
   rsTmp.Close

   'Modified by Morgan 2017/5/3 電子公文
   'If bPrompt = True Then
   If bPrompt = True And frm03020401_01.m_DocNo = "" Then
   'end 2017/5/3
      strTit = "資料檢核"
      strMsg = "與櫃台之來函收文記錄不符, 請確認"
      nResponse = MsgBox(strMsg, vbOKCancel, strTit)
      If nResponse = vbCancel Then
         PromptIfTaiwanNoResult = False
      End If
   End If
   Set rsTmp = Nothing
End Function
'2006/6/1 END

'add by nickc 2006/07/21
Function ChkTmData() As Boolean
ChkTmData = False
Dim rsTmp1 As New ADODB.Recordset
Is308Monther = False
IsHaveTM15 = False
strSql = "select * from trademark,caseprogress where cp09='" & m_CP09 & "' and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp31 is null and cp10='308' "
Set rsTmp1 = New ADODB.Recordset
With rsTmp1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 Then
        Is308Monther = True
        If CheckStr(.Fields("tm15")) <> "" Then
            IsHaveTM15 = True
        End If
        If IsHaveTM15 = True Then
            If CheckStr(.Fields("tm21")) = "" Or CheckStr(.Fields("tm22")) = "" Then
                MsgBox "母案專用期間資料不正確，無法進行下一步，請先補母案專用期間資料！", , "錯誤！"
                Exit Function
            End If
            If CheckStr(.Fields("tm14")) = "" Then
                MsgBox "母案公告日資料不正確，無法進行下一步，請先補母案公告日資料！", , "錯誤！"
                Exit Function
            End If
        Else
            If CheckStr(.Fields("tm11")) = "" Then
                MsgBox "母案申請日資料不正確，無法進行下一步，請先補母案申請日資料！", , "錯誤！"
                Exit Function
            End If
        End If
    End If
End With
Set rsTmp1 = Nothing
ChkTmData = True
End Function
