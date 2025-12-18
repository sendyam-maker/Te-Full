VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010012_03 
   BorderStyle     =   1  '單線固定
   Caption         =   "內部收文"
   ClientHeight    =   5724
   ClientLeft      =   4572
   ClientTop       =   852
   ClientWidth     =   9300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5724
   ScaleWidth      =   9300
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5640
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   945
      Width           =   3492
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   945
      Width           =   3132
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   645
      Width           =   3132
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5640
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   645
      Width           =   3492
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   8340
      TabIndex        =   1
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7515
      TabIndex        =   0
      Top             =   60
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   3132
      Left            =   48
      TabIndex        =   18
      Top             =   2520
      Width           =   9132
      _ExtentX        =   16108
      _ExtentY        =   5525
      _Version        =   393216
      AllowUserResizing=   3
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
      Left            =   1380
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2145
      Width           =   7752
      VariousPropertyBits=   671107103
      ForeColor       =   -2147483641
      Size            =   "13679;476"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM07 
      Height          =   270
      Left            =   1380
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1845
      Width           =   7752
      VariousPropertyBits=   671107103
      Size            =   "13679;476"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM06 
      Height          =   270
      Left            =   1380
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1545
      Width           =   7755
      VariousPropertyBits=   671107103
      Size            =   "13679;476"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM05 
      Height          =   270
      Left            =   1380
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1245
      Width           =   7755
      VariousPropertyBits=   671107103
      Size            =   "13679;476"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家 :"
      Height          =   255
      Index           =   8
      Left            =   4680
      TabIndex        =   17
      Top             =   945
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   255
      Left            =   60
      TabIndex        =   16
      Top             =   2145
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "案件日文名稱 :"
      Height          =   255
      Left            =   60
      TabIndex        =   15
      Top             =   1845
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "案件英文名稱 :"
      Height          =   255
      Left            =   60
      TabIndex        =   14
      Top             =   1545
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "案件中文名稱 :"
      Height          =   255
      Left            =   60
      TabIndex        =   13
      Top             =   1245
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "申請案號 :"
      Height          =   255
      Index           =   13
      Left            =   60
      TabIndex        =   12
      Top             =   945
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   11
      Top             =   645
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "審定號 :"
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   10
      Top             =   645
      Width           =   735
   End
End
Attribute VB_Name = "frm010012_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/16 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Morgan 2021/5/11 改成Form2.0 ( textTM05,grdList...)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo by Morgan2010/12/27 申請案號欄已修改
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/22 日期欄已修改
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
'Modified by Lydia 2020/04/21 改為Form型態
'Dim m_PrevForm As String
Dim m_PrevForm As Form

Private Sub cmdCancel_Click()
   
   'Modified by Lydia 2020/04/21 改為Form型態
   'Select Case m_PrevForm
   '   Case "frm010012_01":
   '      frm010012_01.Show
   '   Case "frm010012_02":
   '      frm010012_02.Show
   '   Case "frm010012_04":
   '      frm010012_04.Show
   '   Case "frm010012_05":
   '      frm010012_05.Show
   '   Case "frm010012_06":
   '      frm010012_06.Show
   '   Case "frm010012_07":
   '      frm010012_07.Show
   'End Select
   If UCase(TypeName(m_PrevForm)) <> "NOTHING" Then
       m_PrevForm.Show
   End If
   'end 2020/04/21
   Unload Me
End Sub

Private Sub cmdok_Click()
   'Modified by Lydia 2020/04/21 改為Form型態
   'Select Case m_PrevForm
   '   Case "frm010012_01":
   '      frm010012_01.SetData m_CP09, 6, False
   '      frm010012_01.Show
   '   Case "frm010012_02":
   '      frm010012_02.SetData m_CP09, 6, False
   '      frm010012_02.Show
   '   Case "frm010012_04":
   '      frm010012_04.SetData m_CP09, 6, False
   '      frm010012_04.Show
   '   Case "frm010012_05":
   '      frm010012_05.SetData m_CP09, 6, False
   '      frm010012_05.Show
   '   Case "frm010012_06":
   '      frm010012_06.SetData m_CP09, 6, False
   '      frm010012_06.Show
   '   Case "frm010012_07":
   '      frm010012_07.SetData m_CP09, 6, False
   '      frm010012_07.Show
   'End Select
   If UCase(TypeName(m_PrevForm)) <> "NOTHING" Then
       m_PrevForm.SetData m_CP09, 6, False
       m_PrevForm.Show
   End If
   'end 2020/04/21
   
   Unload Me
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
      'm_PrevForm = Empty 'Remove by Lydia 2020/04/21
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
'Modified by Lydia 2020/04/21 改為Form型態
'Public Sub SetParent(ByVal strPrevForm As String)
'   m_PrevForm = strPrevForm
'End Sub
Public Sub SetParent(ByRef strPrevForm As Form)
   Set m_PrevForm = strPrevForm
End Sub
'end 2020/04/21

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
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Private Sub QueryPatent()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   m_TM10 = Empty
   
   ' 取得商標基本檔的相關項目
   strSql = "SELECT * FROM PATENT " & _
            "WHERE PA01 = '" & m_TM01 & "' AND " & _
                  "PA02 = '" & m_TM02 & "' AND " & _
                  "PA03 = '" & m_TM03 & "' AND " & _
                  "PA04 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
      ' 申請國家
      If IsNull(rsTmp.Fields("PA09")) = False Then
         m_TM10 = rsTmp.Fields("PA09")
         textTM10 = GetNationName(rsTmp.Fields("PA09"), 0)
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("PA11")) = False Then
         textTM12 = rsTmp.Fields("PA11")
      End If
      ' 審定號
      textTM15 = Empty
      ' 商標名稱(中)
      If IsNull(rsTmp.Fields("PA05")) = False Then
         textTM05 = rsTmp.Fields("PA05")
      End If
      ' 商標名稱(英)
      If IsNull(rsTmp.Fields("PA06")) = False Then
         textTM06 = rsTmp.Fields("PA06")
      End If
      ' 商標名稱(日)
      If IsNull(rsTmp.Fields("PA07")) = False Then
         textTM07 = rsTmp.Fields("PA07")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("PA26")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("PA26"), 0)
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Private Sub QueryServicePractice()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   m_TM10 = Empty
   
   ' 取得商標基本檔的相關項目
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_TM01 & "' AND " & _
                  "SP02 = '" & m_TM02 & "' AND " & _
                  "SP03 = '" & m_TM03 & "' AND " & _
                  "SP04 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_TM10 = rsTmp.Fields("SP09")
         textTM10 = GetNationName(rsTmp.Fields("SP09"), 0)
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("SP11")) = False Then
         textTM12 = rsTmp.Fields("SP11")
      End If
      ' 商標名稱(中)
      If IsNull(rsTmp.Fields("SP05")) = False Then
         textTM05 = rsTmp.Fields("SP05")
      End If
      ' 商標名稱(英)
      If IsNull(rsTmp.Fields("SP06")) = False Then
         textTM06 = rsTmp.Fields("SP06")
      End If
      ' 商標名稱(日)
      If IsNull(rsTmp.Fields("SP07")) = False Then
         textTM07 = rsTmp.Fields("SP07")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("SP08")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("SP08"), 0)
      End If
   End If
   rsTmp.Close
End Sub

' 查詢資料庫取得資料
Public Sub QueryData()
   ' 依系統別來取得基本檔的欄位內容
   Select Case m_TM01
      Case "T", "TF", "CFT", "FCT":
         QueryTradeMark
      Case "P", "CFP", "FCP":
         QueryPatent
      'add by sonia 2019/8/7
      Case "LA":
         QueryHireCase
      Case "L", "CFL", "FCL", "LIN", "ACS":
         QueryLawCase
      'end 2019/8/7
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
   Dim bSpecial As Boolean
   
   InitialGrdList
   
   m_CP09 = Empty

   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP01 = '" & m_TM01 & "' AND " & _
                  "CP02 = '" & m_TM02 & "' AND " & _
                  "CP03 = '" & m_TM03 & "' AND " & _
                  "CP04 = '" & m_TM04 & "' "
   'Added by Lydia 2019/06/12 收文日降冪+CREATE DATE降冪+CREATE TIME降冪排序顯示(配合FCP勘誤作業)
   strSql = strSql & "ORDER BY CP05 DESC, CP66 DESC, CP67 DESC "
   
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
            grdList.TextMatrix(grdList.row, 1) = rsTmp.Fields("CP09")
         End If
         ' 收文日
         If IsNull(rsTmp.Fields("CP05")) = False Then
            grdList.TextMatrix(grdList.row, 2) = TAIWANDATE(rsTmp.Fields("CP05"))
         End If
         ' 案件性質
         If IsNull(rsTmp.Fields("CP10")) = False Then
            grdList.TextMatrix(grdList.row, 3) = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         End If
         ' 發文日
         If IsNull(rsTmp.Fields("CP27")) = False Then
            grdList.TextMatrix(grdList.row, 4) = TAIWANDATE(rsTmp.Fields("CP27"))
         End If
         ' 結果
         If IsNull(rsTmp.Fields("CP24")) = False Then
            grdList.TextMatrix(grdList.row, 5) = rsTmp.Fields("CP24")
         End If
         ' 後金
         If IsNull(rsTmp.Fields("CP19")) = False Then
            grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("CP19")
            If IsEmptyText(rsTmp.Fields("CP19")) = False Then
               bSpecial = True
            End If
         End If
         ' 相關人
         bDeal = False
        'Modify By Cheng 2004/05/12
        '抓對造名稱非對造案件名稱
'         If bDeal = False And IsNull(rsTmp.Fields("CP37")) = False Then
'            If IsEmptyText(rsTmp.Fields("CP37")) = False Then
'               grdList.TextMatrix(grdList.Row, 7) = rsTmp.Fields("CP37")
'               bDeal = True
'            End If
'         End If
'         If bDeal = False And IsNull(rsTmp.Fields("CP38")) = False Then
'            If IsEmptyText(rsTmp.Fields("CP38")) = False Then
'               grdList.TextMatrix(grdList.Row, 7) = rsTmp.Fields("CP38")
'               bDeal = True
'            End If
'         End If
'         If bDeal = False And IsNull(rsTmp.Fields("CP39")) = False Then
'            If IsEmptyText(rsTmp.Fields("CP39")) = False Then
'               grdList.TextMatrix(grdList.Row, 7) = rsTmp.Fields("CP39")
'               bDeal = True
'            End If
'         End If
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
        'End
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
         If bDeal = False Then: grdList.Text = Empty
         ' 特殊欄位
         If bSpecial = True Then
            grdList.TextMatrix(grdList.row, 8) = "1"
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
   'Add By Cheng 2002/07/18
   Set frm010012_03 = Nothing
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
      If grdList.CellBackColor <> &H80000005 Then
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

'add by sonia 2019/8/7
Private Sub QueryHireCase()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   m_TM10 = Empty
   
   ' 取得商標基本檔的相關項目
   strSql = "SELECT * FROM Hirecase " & _
            "WHERE HC01 = '" & m_TM01 & "' AND " & _
                  "HC02 = '" & m_TM02 & "' AND " & _
                  "HC03 = '" & m_TM03 & "' AND " & _
                  "HC04 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
      ' 申請國家
      m_TM10 = "000"
      textTM10 = GetNationName("000", 0)
      ' 案件名稱
      If IsNull(rsTmp.Fields("HC06")) = False Then
         textTM05 = rsTmp.Fields("HC06")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("HC05")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("HC05"), 0)
      End If
   End If
   rsTmp.Close
End Sub

Private Sub QueryLawCase()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   m_TM10 = Empty
   
   ' 取得商標基本檔的相關項目
   strSql = "SELECT * FROM Lawcase " & _
            "WHERE LC01 = '" & m_TM01 & "' AND " & _
                  "LC02 = '" & m_TM02 & "' AND " & _
                  "LC03 = '" & m_TM03 & "' AND " & _
                  "LC04 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
      ' 申請國家
      If IsNull(rsTmp.Fields("LC15")) = False Then
         m_TM10 = rsTmp.Fields("LC15")
         textTM10 = GetNationName(rsTmp.Fields("LC15"), 0)
      End If
      ' 案件名稱(中)
      If IsNull(rsTmp.Fields("LC05")) = False Then
         textTM05 = rsTmp.Fields("LC05")
      End If
      ' 案件名稱(英)
      If IsNull(rsTmp.Fields("LC06")) = False Then
         textTM06 = rsTmp.Fields("LC06")
      End If
      ' 案件名稱(日)
      If IsNull(rsTmp.Fields("LC07")) = False Then
         textTM07 = rsTmp.Fields("LC07")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("LC11")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("LC11"), 0)
      End If
   End If
   rsTmp.Close
End Sub
'end 2019/8/7
