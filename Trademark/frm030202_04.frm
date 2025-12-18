VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030202_04 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文"
   ClientHeight    =   5745
   ClientLeft      =   5985
   ClientTop       =   2100
   ClientWidth     =   9345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9345
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   786
      Width           =   2772
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1500
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   786
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   2532
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   2772
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   8280
      TabIndex        =   1
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7320
      TabIndex        =   0
      Top             =   60
      Width           =   912
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   2055
      Left            =   90
      TabIndex        =   26
      Top             =   3630
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   3625
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
   Begin MSForms.TextBox textTM81 
      Height          =   285
      Left            =   1500
      TabIndex        =   25
      Top             =   3240
      Width           =   7695
      VariousPropertyBits=   671105055
      Size            =   "13573;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM80 
      Height          =   285
      Left            =   1500
      TabIndex        =   24
      Top             =   2928
      Width           =   7695
      VariousPropertyBits=   671105055
      Size            =   "13573;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM79 
      Height          =   285
      Left            =   1500
      TabIndex        =   23
      Top             =   2622
      Width           =   7695
      VariousPropertyBits=   671105055
      Size            =   "13573;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM78 
      Height          =   285
      Left            =   1500
      TabIndex        =   22
      Top             =   2316
      Width           =   7695
      VariousPropertyBits=   671105055
      Size            =   "13573;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1500
      TabIndex        =   21
      Top             =   2010
      Width           =   7695
      VariousPropertyBits=   671105055
      Size            =   "13573;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM07 
      Height          =   285
      Left            =   1500
      TabIndex        =   20
      Top             =   1704
      Width           =   7695
      VariousPropertyBits=   671105055
      Size            =   "13573;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM06 
      Height          =   285
      Left            =   1500
      TabIndex        =   19
      Top             =   1398
      Width           =   7695
      VariousPropertyBits=   671105055
      Size            =   "13573;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM05 
      Height          =   285
      Left            =   1500
      TabIndex        =   18
      Top             =   1092
      Width           =   7695
      VariousPropertyBits=   671105055
      Size            =   "13573;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "申請人5 :"
      Height          =   180
      Left            =   180
      TabIndex        =   17
      Top             =   3292
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "申請人4 :"
      Height          =   180
      Left            =   180
      TabIndex        =   16
      Top             =   2980
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "申請人3 :"
      Height          =   180
      Left            =   180
      TabIndex        =   15
      Top             =   2674
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請人2 :"
      Height          =   180
      Left            =   180
      TabIndex        =   14
      Top             =   2368
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請國家 :"
      Height          =   180
      Index           =   8
      Left            =   4800
      TabIndex        =   13
      Top             =   838
      Width           =   810
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "申請人1 :"
      Height          =   180
      Left            =   180
      TabIndex        =   12
      Top             =   2062
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "案件日文名稱 :"
      Height          =   180
      Left            =   180
      TabIndex        =   11
      Top             =   1756
      Width           =   1170
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "案件英文名稱 :"
      Height          =   180
      Left            =   180
      TabIndex        =   10
      Top             =   1450
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "案件中文名稱 :"
      Height          =   180
      Left            =   180
      TabIndex        =   9
      Top             =   1144
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號 :"
      Height          =   180
      Index           =   13
      Left            =   180
      TabIndex        =   8
      Top             =   838
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號 :"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   7
      Top             =   532
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "審定號 :"
      Height          =   180
      Index           =   1
      Left            =   4800
      TabIndex        =   6
      Top             =   532
      Width           =   630
   End
End
Attribute VB_Name = "frm030202_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/03/18 grdList : MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉
'Memo by Lydia 2021/09/01 改成Form2.0 ; textTM05、textTM06、textTM07、textTM23、textTM78、textTM79、textTM80、textTM81、grdList改字型=新細明體-ExtB
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
' 不列出的收文號
Dim m_NOCP09 As String
'
Dim m_CurrSel As Integer
' 前一畫面
Dim m_PrevForm As Form

Private Sub cmdCancel_Click()
   m_PrevForm.Show
   Unload Me
End Sub

Private Sub cmdOK_Click()
   Unload Me
   If IsObject(m_PrevForm) = True Then
      m_PrevForm.SetData 99, m_CP09, False
      m_PrevForm.Show
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
   'add by nickc 2007/02/01
   textTM78.BackColor = &H8000000F
   textTM79.BackColor = &H8000000F
   textTM80.BackColor = &H8000000F
   textTM81.BackColor = &H8000000F
   
   MoveFormToCenter Me
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_NOCP09 = Empty
      
      Set m_PrevForm = Nothing
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
      Case 4: m_NOCP09 = strData
   End Select
End Sub

Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

' 取得國家的名稱
Private Function GetNation(ByVal strNation As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   GetNation = Empty
   If IsEmptyText(strNation) = False Then
      strSql = "SELECT * FROM NATION " & _
               "WHERE NA01 = '" & strNation & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("NA03")) = False Then
            GetNation = rsTmp.Fields("NA03")
         End If
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Function

' 由客戶代碼取得客戶名稱
Private Function GetCustomer(ByVal strData As String) As String
   Dim rsTmp As ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   
   GetCustomer = Empty
   If IsEmptyText(strData) = False Then
      Set rsTmp = New ADODB.Recordset
      If Len(strData) > 8 Then
         strSql = "SELECT * FROM Customer " & _
                  "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                        "CU02 = '" & Mid(strData, 9, 1) & "'"
      Else
         strSql = "SELECT * FROM Customer " & _
                  "WHERE CU01 = '" & Mid(strData, 1, 8) & "'"
      End If
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("CU04")) = False Then
            GetCustomer = rsTmp.Fields("CU04")
         ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
            GetCustomer = rsTmp.Fields("CU05")
         ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
            GetCustomer = rsTmp.Fields("CU06")
         End If
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Function

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
      textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
         textTM10 = GetNation(rsTmp.Fields("TM10"))
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
         textTM23 = GetCustomer(rsTmp.Fields("TM23"))
      End If
      'add by nickc 2007/02/01
      If IsNull(rsTmp.Fields("TM78")) = False Then
         textTM78 = GetCustomer(rsTmp.Fields("TM78"))
      End If
      If IsNull(rsTmp.Fields("TM79")) = False Then
         textTM79 = GetCustomer(rsTmp.Fields("TM79"))
      End If
      If IsNull(rsTmp.Fields("TM80")) = False Then
         textTM80 = GetCustomer(rsTmp.Fields("TM80"))
      End If
      If IsNull(rsTmp.Fields("TM81")) = False Then
         textTM81 = GetCustomer(rsTmp.Fields("TM81"))
      End If
   End If
   rsTmp.Close
   ' 顯示符合條件的資料
   ListData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 由案件性質代碼取得案件性質名稱
Private Function GetCaseType(ByVal strKey1 As String, ByVal StrKey2 As String) As String
   Dim rsTmp As ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   
   GetCaseType = Empty
   If IsEmptyText(strKey1) = False And IsEmptyText(StrKey2) = False Then
      Set rsTmp = New ADODB.Recordset
      strSql = "SELECT * FROM CasePropertyMap " & _
               "WHERE CPM01 = '" & strKey1 & "' AND " & _
                     "CPM02 = '" & StrKey2 & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If m_TM10 < "010" Then
            If IsNull(rsTmp.Fields("CPM03")) = False Then
               GetCaseType = rsTmp.Fields("CPM03")
            End If
         Else
            If IsNull(rsTmp.Fields("CPM04")) = False Then
               GetCaseType = rsTmp.Fields("CPM04")
            End If
         End If
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Function

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
                  "CP04 = '" & m_TM04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         bSpecial = False
         
         ' 不列出前畫面所使用的發文號資料
         If IsEmptyText(m_NOCP09) = False Then
            If IsNull(rsTmp.Fields("CP09")) = False Then
               If m_NOCP09 = rsTmp.Fields("CP09") Then
                  GoTo NextRecord
               End If
            End If
         End If
         
         ' 列入資料
         grdList.Rows = grdList.Rows + 1
         grdList.row = grdList.Rows - 1
         ' 收文號
         If IsNull(rsTmp.Fields("CP09")) = False Then
            grdList.col = 1
            grdList.Text = rsTmp.Fields("CP09")
         End If
         ' 收文日
         If IsNull(rsTmp.Fields("CP05")) = False Then
            grdList.col = 2
            grdList.Text = ChangeWStringToTString(rsTmp.Fields("CP05"))
         End If
         ' 案件性質
         If IsNull(rsTmp.Fields("CP10")) = False Then
            grdList.col = 3
            If m_TM10 < "010" Then
               grdList.Text = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
            Else
              grdList.Text = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
            End If
         End If
         ' 發文日
         If IsNull(rsTmp.Fields("CP27")) = False Then
            grdList.col = 4
            grdList.Text = ChangeWStringToTString(rsTmp.Fields("CP27"))
         End If
         ' 後金
         If IsNull(rsTmp.Fields("CP19")) = False Then
            grdList.col = 5
            grdList.Text = rsTmp.Fields("CP19")
            If IsEmptyText(rsTmp.Fields("CP19")) = False Then
               bSpecial = True
            End If
         End If
         ' 相關人
         grdList.col = 7
         bDeal = False
         If bDeal = False And IsNull(rsTmp.Fields("CP37")) = False Then
            If IsEmptyText(rsTmp.Fields("CP37")) = False Then
               grdList.Text = rsTmp.Fields("CP37")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP38")) = False Then
            If IsEmptyText(rsTmp.Fields("CP38")) = False Then
               grdList.Text = rsTmp.Fields("CP38")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP39")) = False Then
            If IsEmptyText(rsTmp.Fields("CP39")) = False Then
               grdList.Text = rsTmp.Fields("CP39")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP50")) = False Then
            If IsEmptyText(rsTmp.Fields("CP50")) = False Then
               grdList.Text = rsTmp.Fields("CP50")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP51")) = False Then
            If IsEmptyText(rsTmp.Fields("CP51")) = False Then
               grdList.Text = rsTmp.Fields("CP51")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP52")) = False Then
            If IsEmptyText(rsTmp.Fields("CP52")) = False Then
               grdList.Text = rsTmp.Fields("CP52")
               bDeal = True
            End If
         End If
         If bDeal = False And IsNull(rsTmp.Fields("CP56")) = False Then
            If IsEmptyText(rsTmp.Fields("CP56")) = False Then
               grdList.Text = GetCustomer(rsTmp.Fields("CP56"))
               bDeal = True
            End If
         End If
         If bDeal = False Then: grdList.Text = Empty
         ' 特殊欄位
         If bSpecial = True Then
            grdList.col = 8
            grdList.Text = "1"
         End If
NextRecord:
         rsTmp.MoveNext
      Loop
      
      'Added by Lydia 2022/03/18 MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉
      If grdList.Rows >= 2 Then
           grdList.FixedRows = 1
      End If
      'end 2022/03/18
   End If
   
   ' 設定第一列為所選取的記錄
   grdList_SetSelection 1

End Sub

' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
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

' 設定Grid List的一列為選取的狀態
Private Sub grdList_SetSelection(ByVal nSel As Integer)
   If nSel > 0 And nSel < grdList.Rows And grdList.Rows >= 2 Then
      grdList.row = nSel
      grdList_SelChange
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/19
   Set frm030202_04 = Nothing
End Sub

Private Sub grdList_SelChange()
   If grdList.Rows > 1 Then
      If grdList.row > 0 Then
         m_CP09 = grdList.TextMatrix(grdList.row, 1)
      End If
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
            If grdList.TextMatrix(grdList.row, 8) = "1" Then
               If grdList.CellBackColor <> &HFF& Then: grdList.CellBackColor = &HFF&
               If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
            Else
               grdList.col = nCol
               If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
               If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
            End If
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

