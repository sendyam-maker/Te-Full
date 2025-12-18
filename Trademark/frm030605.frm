VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030605 
   BorderStyle     =   1  '單線固定
   Caption         =   "國內公報代理人/事務所名稱查詢"
   ClientHeight    =   5745
   ClientLeft      =   180
   ClientTop       =   975
   ClientWidth     =   9330
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "新細明體-ExtB"
      Size            =   9
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9330
   Begin VB.CommandButton cmdQuery 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   7320
      TabIndex        =   1
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   8280
      TabIndex        =   2
      Top             =   60
      Width           =   912
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   4692
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   9132
      _ExtentX        =   16113
      _ExtentY        =   8281
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
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
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.TextBox textTA03 
      Height          =   300
      Left            =   1920
      TabIndex        =   0
      Top             =   570
      Width           =   2175
      VariousPropertyBits=   679493659
      MaxLength       =   12
      Size            =   "3831;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代理人/事務所名稱 :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "frm030605"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/11 Form2.0已修改 textTA03/grdList
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

Dim m_CurrSel As Integer

' 使用者按下離開按紐
Private Sub cmdExit_Click()
   Unload Me
End Sub
' 使用者按下查詢按紐
Private Sub cmdQuery_Click()
   QueryDB
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   InitialGrdList (True) 'Add byAmy 2022/01/11
End Sub

' 查詢資料庫
Private Sub QueryDB()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
    
   grdList.Clear
   'Modify by Amy 2022/01/11 改Form2.0 原Grid不支援UniCode
   ' 設定查詢的與法
   '2015/8/17 MODIFY BY SONIA 加入同時查事務所欄TA04
   strSql = "SELECT ta02 as 代理人編號,ta03 as 代理人名稱,ta04 as 事務所名稱  FROM TAGENT " & _
            "WHERE TA01 = 'T' AND (TA03 LIKE '%" & textTA03 & "%' OR TA04 LIKE '%" & textTA03 & "%') ORDER BY TA02"
   ' 查詢資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount <= 0 Then
      strTit = "查詢代理人資料"
      strMsg = "資料庫內無符合條件的資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      'GoTo EXITSUB
   End If
   Set grdList.Recordset = rsTmp
   If rsTmp.RecordCount <= 0 Then
        InitialGrdList (True)
   Else
        InitialGrdList (False)
   End If
   
   'ListData rsTmp
   
   rsTmp.Close
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 初始化 GridList
Private Sub InitialGrdList(bolShowR2 As Boolean)
   'Modify by Amy 2022/01/11 原Grid 不支援UniCode
   Dim i As Integer
   
   i = 0
   If bolShowR2 = True Then grdList.Rows = 2
   grdList.Cols = 3
   grdList.ColWidth(i) = 300
   grdList.row = 0
   grdList.col = i
   grdList.Text = "代理人編號"
   grdList.ColWidth(i) = 1200
   grdList.ColAlignment(i) = flexAlignCenterCenter
   i = i + 1
   grdList.col = i
   grdList.Text = "代理人名稱"
   grdList.ColWidth(i) = 2000
   grdList.ColAlignment(i) = flexAlignLeftCenter
   i = i + 1
   grdList.col = i
   grdList.Text = "事務所名稱"
   grdList.ColWidth(i) = 3600
   grdList.ColAlignment(i) = flexAlignLeftCenter
   'end 2022/01/11
End Sub

'Mark by Amy 2022/01/11 換成MSHFlexGrid可不需使用
Private Sub ListData(ByRef rsTmp As ADODB.Recordset)
'   Dim nRow As Integer
'
'   If rsTmp.RecordCount > 0 Then
'      rsTmp.MoveFirst
'      Do While rsTmp.EOF = False
'         grdList.Rows = grdList.Rows + 1
'         nRow = grdList.Rows - 1
'         ' 代理人代號
'         If IsNull(rsTmp.Fields("TA02")) = False Then
'            grdList.TextMatrix(nRow, 1) = rsTmp.Fields("TA02")
'         End If
'         ' 代理人姓名
'         If IsNull(rsTmp.Fields("TA03")) = False Then
'            grdList.TextMatrix(nRow, 2) = rsTmp.Fields("TA03")
'         End If
'         ' 事務所名稱
'         If IsNull(rsTmp.Fields("TA04")) = False Then
'            grdList.TextMatrix(nRow, 3) = rsTmp.Fields("TA04")
'         End If
'         rsTmp.MoveNext
'      Loop
'   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/19
   Set frm030605 = Nothing
End Sub

Private Sub grdList_SelChange()
   grdList_ShowSelection
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = grdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = grdList.row Then
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

' 代理人名稱
Private Sub textTA03_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTA03) = False Then
      If StrLength(textTA03) > 12 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代理人名稱內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTA03_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTA03.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

Private Sub textTA03_GotFocus()
   InverseTextBox textTA03
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTA03.IMEMode = 1
   OpenIme
End Sub


