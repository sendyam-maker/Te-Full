VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030102_02 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標申請案號輸入"
   ClientHeight    =   3810
   ClientLeft      =   4830
   ClientTop       =   4020
   ClientWidth     =   9150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   9150
   Begin VB.CheckBox Check1 
      Caption         =   "申請收據"
      Height          =   252
      Index           =   1
      Left            =   2520
      TabIndex        =   6
      Top             =   3420
      Width           =   1185
   End
   Begin VB.CheckBox Check1 
      Caption         =   "申請書"
      Height          =   252
      Index           =   0
      Left            =   1200
      TabIndex        =   5
      Top             =   3420
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7080
      TabIndex        =   8
      Top             =   50
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6180
      TabIndex        =   7
      Top             =   50
      Width           =   852
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8220
      TabIndex        =   9
      Top             =   50
      Width           =   852
   End
   Begin VB.TextBox textTM12 
      Height          =   264
      Left            =   5580
      MaxLength       =   30
      TabIndex        =   1
      Top             =   2340
      Width           =   2532
   End
   Begin VB.TextBox textCP45 
      Height          =   264
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   4
      Top             =   3060
      Width           =   2532
   End
   Begin VB.TextBox textPrint 
      Height          =   264
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   2
      Top             =   2700
      Width           =   372
   End
   Begin VB.TextBox textTM11 
      Height          =   264
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   0
      Top             =   2340
      Width           =   1092
   End
   Begin VB.TextBox textCF09 
      Height          =   264
      Left            =   5100
      MaxLength       =   12
      TabIndex        =   3
      Top             =   2700
      Width           =   1005
   End
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   540
      Width           =   2412
   End
   Begin VB.TextBox textTM09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   395
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1620
      Width           =   6732
   End
   Begin VB.TextBox textTM32 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   699
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1980
      Width           =   6732
   End
   Begin VB.TextBox textTM08 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1260
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   540
      Width           =   2532
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   300
      Left            =   1200
      TabIndex        =   31
      Top             =   900
      Width           =   6735
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "11880;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5700
      TabIndex        =   30
      Top             =   1260
      Width           =   2415
      VariousPropertyBits=   671105055
      Size            =   "4260;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "附件 :"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   29
      Top             =   3420
      Width           =   972
   End
   Begin VB.Label Label3 
      Caption         =   "申請案號 :"
      Height          =   252
      Left            =   4620
      TabIndex        =   28
      Top             =   2340
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "彼所案號 :"
      Height          =   252
      Index           =   9
      Left            =   120
      TabIndex        =   27
      Top             =   3060
      Width           =   972
   End
   Begin VB.Label Label22 
      Caption         =   "列印定稿 :"
      Height          =   252
      Left            =   120
      TabIndex        =   26
      Top             =   2700
      Width           =   972
   End
   Begin VB.Label Label23 
      Caption         =   "(N:不印)"
      Height          =   252
      Left            =   1680
      TabIndex        =   25
      Top             =   2700
      Width           =   852
   End
   Begin VB.Label Label10 
      Caption         =   "申請日 :"
      Height          =   252
      Left            =   120
      TabIndex        =   24
      Top             =   2340
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "大約"
      Height          =   252
      Index           =   12
      Left            =   4620
      TabIndex        =   23
      Top             =   2700
      Width           =   492
   End
   Begin VB.Label Label11 
      Caption         =   "後可接獲審查報告"
      Height          =   255
      Left            =   6210
      TabIndex        =   22
      Top             =   2700
      Width           =   2205
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家 :"
      Height          =   252
      Index           =   8
      Left            =   4620
      TabIndex        =   21
      Top             =   540
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   252
      Index           =   11
      Left            =   4620
      TabIndex        =   19
      Top             =   1260
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "商品類別 :"
      Height          =   252
      Index           =   7
      Left            =   120
      TabIndex        =   18
      Top             =   1620
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "商品組群 :"
      Height          =   252
      Index           =   5
      Left            =   120
      TabIndex        =   17
      Top             =   1980
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "商標種類 :"
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   1260
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   540
      Width           =   852
   End
   Begin VB.Label Label5 
      Caption         =   "案件名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   12
      Top             =   900
      Width           =   972
   End
End
Attribute VB_Name = "frm030102_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/08/10 改成Form2.0 ; cmbTM05、textCP13
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
' 申請國家
Dim m_TM10 As String
' 來函收文日
Dim m_CP05 As String
' 收文號
Dim m_CP09 As String
' 智權人員代號
Dim m_CP13 As String
' 業務區
Dim m_CP12 As String
' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
' 儲存商標基本檔或服務業務基本檔檔案欄位的串列
Dim m_TMSPList() As FIELDITEM
Dim m_TMSPCount As Integer
Dim m_CP64 As String, m_CP10 As String
Dim m_CP27 As String   'add by sonia 2018/2/6
'Add By Sindy 2023/4/24
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2023/4/24 END


'Add By Sindy 2023/4/24
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdCancel_Click()
   Unload Me
   frm030102_01.Show
End Sub

Private Sub cmdExit_Click()
   Unload frm030102_01
   Unload Me
   'frm030102_01.Show
End Sub

Private Sub cmdOK_Click()
   If CheckDataValid = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
   
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 更新欄位輸入的內容
      OnUpdateField
      ' 存檔
      'edit by nick 2004/11/03
      'OnSaveData
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub

      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      'Add By Sindy 2023/4/24
      If Me.m_strIR01 <> "" Then
         Unload frm030102_01
         If Not m_PrevForm Is Nothing Then
            Call m_PrevForm.GoNext
         End If
         Unload Me
      Else
      '2023/4/24 END
         frm030102_01.Show
         '910718 Sieg
         frm030102_01.textTM01 = ""
         frm030102_01.textTM02 = ""
         frm030102_01.textTM02_2 = ""
         frm030102_01.textTM03 = ""
         frm030102_01.textTM04 = ""
         Unload Me
      End If
   End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM08.BackColor = &H8000000F
   textTM09.BackColor = &H8000000F
   textTM10.BackColor = &H8000000F
   textTM32.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   
   MoveFormToCenter Me
   
   'Add By Sindy 2023/4/24
   m_strIR01 = frm030102_01.m_strIR01
   m_strIR02 = frm030102_01.m_strIR02
   m_strIR03 = frm030102_01.m_strIR03
   m_strIR04 = frm030102_01.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2023/4/24 END
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP09 = Empty
   End If
   
   Select Case nType
      ' 本所案號
      Case 0: m_TM01 = strData
      Case 1: m_TM02 = strData
      Case 2: m_TM03 = strData
      Case 3: m_TM04 = strData
      Case 4: m_CP05 = strData
   End Select
End Sub

' 清除商標基本檔檔案欄位串列
Private Sub ClearTMSPFieldList()
   If m_TMSPCount > 0 Then
      Erase m_TMSPList
   End If
   m_TMSPCount = 0
End Sub

' 設定商標基本檔或服務業務基本檔欄位串列中的欄位內容
Private Sub SetTMSPFieldOldData(ByVal strFieldName As String, ByVal strFieldData As String, ByVal nFieldType As Integer)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_TMSPCount - 1
      If m_TMSPList(nPos).fiName = strFieldName Then
         bFind = True
         m_TMSPList(nPos).fiOldData = strFieldData
         m_TMSPList(nPos).fiNewData = strFieldData
         m_TMSPList(nPos).fiType = nFieldType
         Exit For
      End If
   Next nPos
   If bFind = False Then
      ReDim Preserve m_TMSPList(m_TMSPCount + 1)
      m_TMSPList(m_TMSPCount).fiName = strFieldName
      m_TMSPList(m_TMSPCount).fiOldData = strFieldData
      m_TMSPList(m_TMSPCount).fiNewData = strFieldData
      m_TMSPList(m_TMSPCount).fiType = nFieldType
      m_TMSPCount = m_TMSPCount + 1
   End If
End Sub

' 設定商標基本檔或服務業務基本檔欄位串列中的欄位內容
Private Sub SetTMSPFieldNewData(ByVal strFieldName As String, ByVal strFieldData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_TMSPCount - 1
      If m_TMSPList(nPos).fiName = strFieldName Then
         bFind = True
         m_TMSPList(nPos).fiNewData = strFieldData
         Exit For
      End If
   Next nPos
End Sub

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
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
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

' 取得商標基本檔的欄位內容
Private Sub QueryTradeMark()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strTM11 As String
   Dim strTM12 As String
   
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 案件中文名稱
      If IsNull(rsTmp.Fields("TM05")) = False Then: cmbTM05.AddItem rsTmp.Fields("TM05")
      ' 案件英文名稱
      If IsNull(rsTmp.Fields("TM06")) = False Then: cmbTM05.AddItem rsTmp.Fields("TM06")
      ' 案件日文名稱
      If IsNull(rsTmp.Fields("TM07")) = False Then: cmbTM05.AddItem rsTmp.Fields("TM07")
      ' 顯示案件名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 商標種類
      textTM08 = Empty
      If IsNull(rsTmp.Fields("TM08")) = False Then: textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
      ' 商品類別
      textTM09 = Empty
      If IsNull(rsTmp.Fields("TM09")) = False Then: textTM09 = rsTmp.Fields("TM09")
      ' 商品組群
      textTM32 = Empty
      If IsNull(rsTmp.Fields("TM32")) = False Then: textTM32 = rsTmp.Fields("TM32")
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
         textTM10 = GetNationName(rsTmp.Fields("TM10"), 0)
      End If
      ' 申請日
      strTM11 = Empty
      
      '910718 Sieg
      textTM11 = ""
      If IsNull(rsTmp.Fields("TM11")) = False Then
         strTM11 = rsTmp.Fields("TM11")
         textTM11 = TransDate(strTM11, 2)
      End If
      SetTMSPFieldOldData "TM11", strTM11, 1
      ' 申請案號
      strTM12 = Empty
      textTM12 = ""
      If IsNull(rsTmp.Fields("TM12")) = False Then
         strTM12 = rsTmp.Fields("TM12")
         textTM12 = strTM12
      End If
      SetTMSPFieldOldData "TM12", strTM12, 0
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 取得服務業務基本檔的欄位內容
Private Sub QueryServicePractice()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSP10 As String
   Dim strSP11 As String
   
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_TM01 & "' AND " & _
                  "SP02 = '" & m_TM02 & "' AND " & _
                  "SP03 = '" & m_TM03 & "' AND " & _
                  "SP04 = '" & m_TM04 & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ' 案件中文名稱
      If IsNull(rsTmp.Fields("SP05")) = False Then: cmbTM05.AddItem rsTmp.Fields("SP05")
      ' 案件英文名稱
      If IsNull(rsTmp.Fields("SP06")) = False Then: cmbTM05.AddItem rsTmp.Fields("SP06")
      ' 案件日文名稱
      If IsNull(rsTmp.Fields("SP07")) = False Then: cmbTM05.AddItem rsTmp.Fields("SP07")
      ' 顯示案件名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_TM10 = rsTmp.Fields("SP09")
         textTM10 = GetNationName(rsTmp.Fields("SP09"), 0)
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("SP11")) = False Then: textTM12 = rsTmp.Fields("SP11")
      ' 申請日
      strSP10 = Empty
      If IsNull(rsTmp.Fields("SP10")) = False Then: strSP10 = rsTmp.Fields("SP10")
      SetTMSPFieldOldData "SP10", strSP10, 1
      ' 申請案號
      strSP11 = Empty
      If IsNull(rsTmp.Fields("SP11")) = False Then: strSP11 = rsTmp.Fields("SP11")
      SetTMSPFieldOldData "SP11", strSP11, 0
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
    'Add By Cheng 2004/02/2
    strSql = ""
    '93.12.1 恢復 BY SONIA  CFT-009900
    ''edit by nick 2004/09/01 不限制國家
    ''Select Case m_TM10
    ''Case "014" '新加坡
    '    StrSql = " And (CP10='101' Or CP10='107') "
    ''Case Else '其他
    ''    strSQL = " And CP10='101' "
    ''End Select
    Select Case m_TM10
    Case "014" '新加坡
       strSql = " And (CP10='101' Or CP10='107') "
    Case Else '其他
        '2006/1/26 MODIFY BY SONIA
        'strSQL = " And CP10='101' "
        Select Case m_TM01
           Case "CFC"
              strSql = " AND CP10='806' "
           Case Else
              strSql = " And CP10='101' "
        End Select
        '2006/1/26 END
    End Select
    '93.12.1 end
    'End
   'Modify By Cheng 2002/07/11
   '原判斷CP31='Y', 現改為CP10='101'
'   strSQL = "SELECT * FROM CaseProgress " & _
'            "WHERE CP01 = '" & m_TM01 & "' AND " & _
'                  "CP02 = '" & m_TM02 & "' AND " & _
'                  "CP03 = '" & m_TM03 & "' AND " & _
'                  "CP04 = '" & m_TM04 & "' AND " & _
'                  "CP31 = '" & "Y" & "' "
    'Modify By Cheng 2004/02/20
'   strSQL = "SELECT * FROM CaseProgress " & _
'            "WHERE CP01 = '" & m_TM01 & "' AND " & _
'                  "CP02 = '" & m_TM02 & "' AND " & _
'                  "CP03 = '" & m_TM03 & "' AND " & _
'                  "CP04 = '" & m_TM04 & "' AND " & _
'                  "CP10 = '101' "
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP01 = '" & m_TM01 & "' AND " & _
                  "CP02 = '" & m_TM02 & "' AND " & _
                  "CP03 = '" & m_TM03 & "' AND " & _
                  "CP04 = '" & m_TM04 & "' " & strSql
    'End
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 收文號
      If IsNull(rsTmp.Fields("CP09")) = False Then
         m_CP09 = rsTmp.Fields("CP09")
      End If
      ' 業務區
      If IsNull(rsTmp.Fields("CP12")) = False Then
         m_CP12 = rsTmp.Fields("CP12")
      End If
      ' 智權人員代號
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"), True)
      End If
      ' 彼所案號
      If IsNull(rsTmp.Fields("CP45")) = False Then
         textCP45 = rsTmp.Fields("CP45")
      End If
      'Add By Sindy 2018/1/25
      ' 案件性質
      m_CP10 = ""
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
      End If
      ' 進度備註
      m_CP64 = ""
      If IsNull(rsTmp.Fields("CP64")) = False Then
         m_CP64 = rsTmp.Fields("CP64")
      End If
      '2018/1/25 END
      ' 發文日  add by sonia 2018/2/6
      m_CP27 = 0
      If IsNull(rsTmp.Fields("CP27")) = False Then
         m_CP27 = rsTmp.Fields("CP27")
      End If
      'end 2018/2/6
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 讀取資料庫
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
   
   ' 先清除商標基本檔或服務業務基本檔欄位串列
   ClearTMSPFieldList
       
'   ' 讀取案件進度檔
'   QueryCaseProgress
   Select Case m_TM01
      ' 系統類別為CFT的為讀取商標基本檔
      Case "CFT":
         QueryTradeMark
      Case Else:
         QueryServicePractice
   End Select
   ' 讀取案件進度檔
   QueryCaseProgress
   
   ' 大約?可接獲回音(欄位) 以系統別+申請國家代碼+案件性質(通知申請案號)取得案件收費表的回音欄位
   textCF09 = Empty
   strSql = "SELECT * FROM CaseFee " & _
            "WHERE CF01 = '" & m_TM01 & "' AND " & _
                  "CF02 = '" & m_TM10 & "' AND " & _
                  "CF03 = '1101' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CF09")) = False Then
         textCF09 = rsTmp.Fields("CF09")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
    'Add By Cheng 2003/01/09
    '若無回音, 則預設"個月"
    Me.textCF09.Text = "個月"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/19
   Set frm030102_02 = Nothing
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

' 更新欄位的內容
Private Sub OnUpdateField()
   Select Case m_TM01
      ' 系統類別為CFT的為更新商標基本檔
      Case "CFT":
         SetTMSPFieldNewData "TM11", DBDATE(textTM11)
         SetTMSPFieldNewData "TM12", textTM12
      Case Else:
         SetTMSPFieldNewData "SP10", DBDATE(textTM11)
         SetTMSPFieldNewData "SP11", textTM12
   End Select
End Sub

' 更新商標基本檔的相關欄位
Private Sub OnUpdateTradeMark()
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
      
   ' 更新案件進度檔
   strSql = "UPDATE TradeMark SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_TMSPCount - 1
      strTmp = Empty
      If m_TMSPList(nIndex).fiOldData <> m_TMSPList(nIndex).fiNewData Then
         If m_TMSPList(nIndex).fiType = 0 Then
            ' 91.03.25 modify by louis (單引號)
            'strTmp = m_TMSPList(nIndex).fiName & " = '" & m_TMSPList(nIndex).fiNewData & "'"
            strTmp = m_TMSPList(nIndex).fiName & " = '" & ChgSQL(m_TMSPList(nIndex).fiNewData) & "'"
         Else
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = " & 0
            Else
               strTmp = m_TMSPList(nIndex).fiName & " = " & m_TMSPList(nIndex).fiNewData
            End If
         End If
      End If
      If strTmp <> Empty Then
         bDifference = True
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   ' 設定SQL語法更新的條件
   strSql = strSql & " " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "' "
   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
End Sub

' 更新服務業務基本檔的相關欄位
Private Sub OnUpdateServicePractice()
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
      
   ' 更新案件進度檔
   strSql = "UPDATE ServicePractice SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_TMSPCount - 1
      strTmp = Empty
      If m_TMSPList(nIndex).fiOldData <> m_TMSPList(nIndex).fiNewData Then
         If m_TMSPList(nIndex).fiType = 0 Then
            ' 91.03.25 modify by louis (單引號)
            'strTmp = m_TMSPList(nIndex).fiName & " = '" & m_TMSPList(nIndex).fiNewData & "'"
            strTmp = m_TMSPList(nIndex).fiName & " = '" & ChgSQL(m_TMSPList(nIndex).fiNewData) & "'"
         Else
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = " & 0
            Else
               strTmp = m_TMSPList(nIndex).fiName & " = " & m_TMSPList(nIndex).fiNewData
            End If
         End If
      End If
      If strTmp <> Empty Then
         bDifference = True
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   ' 設定SQL語法更新的條件
   strSql = strSql & " " & _
                  "WHERE SP01 = '" & m_TM01 & "' AND " & _
                        "SP02 = '" & m_TM02 & "' AND " & _
                        "SP03 = '" & m_TM03 & "' AND " & _
                        "SP04 = '" & m_TM04 & "' "
   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
End Sub

'edit b nick 2004/11/03
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim nIndex As Integer
   Dim strCP05 As String
   Dim strCP09 As String
   Dim strCP10 As String
   'Dim strCP12 As String
   Dim strCP27 As String
   'Add By Cheng 2002/06/07
   Dim strNP08 As String '下一程序的本所期限
   Dim strNP09 As String '下一程序的法定期限
   Dim strNP22 As String '下一程序的序號
   Dim bInsert As Boolean
   Dim strNP07 As String, strYear As String, strStartUpDay As String 'Add By Sindy 2019/10/16
   
'911106 nick transation
On Error GoTo CheckingErr

   OnSaveData = True
   
cnnConnection.BeginTrans
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新案件進度檔的代理人提申日及彼所案號
   strSql = "UPDATE CaseProgress SET CP47 = " & DBDATE(textTM11) & ", " & _
                                    "CP45 = '" & ChgSQL(textCP45) & "' " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   cnnConnection.Execute strSql
      'add by nick 2005/01/04 更新相同本所案號之相同代理人的彼所案號，若是彼所案號空的話
      If textCP45 <> "" Then
         strSql = "update caseprogress set cp45=" & CNULL(ChgSQL(textCP45)) & " where cp09 in (select cp09 from caseprogress where cp45 is null and CP01 = '" & m_TM01 & "' AND  CP02 = '" & m_TM02 & "' AND CP03 = '" & m_TM03 & "' AND CP04 = '" & m_TM04 & "' and cp09<'C' AND cp44 in (select cp44 from caseprogress where cp09='" & m_CP09 & "' ))"
         cnnConnection.Execute strSql
       End If

   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新基本檔
   Select Case m_TM01
      ' 系統類別為CFT的為更新商標基本檔
      Case "CFT":
         OnUpdateTradeMark
      Case Else:
         OnUpdateServicePractice
   End Select

   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 新增資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   ' 收文日
   strCP05 = DBDATE(m_CP05)
   ' 案件性質
   strCP10 = "1101"
   ' 業務區別 91.8.26 modify by sonia
   'strCP12 = GetStaffDepartment(m_CP13)
   ' 發文日
   strCP27 = DBDATE(SystemDate())
   ' 組成SQL語法
   '2011/12/15 modify by sonia 智權人員改抓PUB_GetAKindSalesNo
   'strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43) " & _
            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & "," & _
                     "'" & strCP09 & "','" & StrCP10 & "','" & m_CP12 & "','" & m_CP13 & "','" & strUserNum & "'," & _
                     "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "') "
   strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43) " & _
            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & "," & _
                     "'" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "'," & _
                     "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "') "
   cnnConnection.Execute strSql

    'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
    Pub_UpdateFromMaxCP27 m_TM01, m_TM02, m_TM03, m_TM04

   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新下一程序檔其下一程序為收達或提申的資料更新其是否續辦欄位為Y
   strSql = "UPDATE NextProgress SET NP06 = 'Y' " & _
            "WHERE NP01 = '" & m_CP09 & "' AND " & _
                  "NP02 = '" & m_TM01 & "' AND " & _
                  "NP03 = '" & m_TM02 & "' AND " & _
                  "NP04 = '" & m_TM03 & "' AND " & _
                  "NP05 = '" & m_TM04 & "' AND " & _
                  "(NP07 = 997 OR NP07 = 998) "
   cnnConnection.Execute strSql
   
   'Add By Cheng 2002/06/07
   '若申請國家為"菲律賓"時, 新增資料至下一程序檔, 下一程序為"使用宣誓", 法定期限為申請日 + 3年, 本所期限 = 法定期限 - 半年(2007/07/11 改為-2個月)
   'MODIFY BY SONIA 2014/5/29 加入波多黎各112
   'MODIFY BY Sindy 2019/10/16 + 302.奈及利亞
   'modify by sonia 2022/10/21 +318莫三比克,申請日+5年
   '**********加入國家時frm030403的Process之可辦期限也要改
   If m_TM10 = "030" Or m_TM10 = "112" Or m_TM10 = "302" Or m_TM10 = "318" Then
      'Add By Sindy 2019/10/16
      If m_TM10 = "302" Then
         strNP07 = "109" '緩審延展
         If ClsPDGetNationTax(10, "302", strStartUpDay, strYear) = True Then
            '法定期限=申請日+國家檔NA13商標專用年度
            If Val(strYear) > 0 Then
               strNP09 = DBDATE(DateAdd("yyyy", Val(strYear), ChangeWStringToWDateString(DBDATE(textTM11.Text))))
            End If
            '本所期限=法定期限-2個月
            strNP08 = DBDATE(DateAdd("m", -2, ChangeWStringToWDateString(DBDATE(strNP09))))
         End If
         If strNP09 = "" Then strNP09 = strSrvDate(1)
         If strNP08 = "" Then strNP08 = strSrvDate(1)
      Else
      '2019/10/16 END
         strNP07 = "105" '使用宣誓
         '法定期限
           'Modify By Cheng 2003/09/02
   '      strNP09 = DBDATE(Format(DateSerial(Val(DBYEAR(textTM11.Text)) + 3, Val(DBMONTH(textTM11.Text)), Val(DBDAY(textTM11.Text)))))
         strNP09 = DBDATE(DateAdd("yyyy", 3, ChangeWStringToWDateString(DBDATE(textTM11.Text))))
         'add by sonia 2022/10/21  318莫三比克,申請日+5年
         If m_TM10 = "318" Then
            strNP09 = DBDATE(DateAdd("yyyy", 5, ChangeWStringToWDateString(DBDATE(textTM11.Text))))
         End If
         'end 2022/10/21
         '本所期限
           'Modify By Cheng 2003/09/02
   '      strNP08 = DBDATE(Format(DateSerial(Val(DBYEAR(strNP09)), Val(DBMONTH(strNP09)) - 6, Val(DBDAY(strNP09)))))
         'edit by nickc 2007/07/11 改為-2個月
         'strNP08 = DBDATE(DateAdd("m", -6, ChangeWStringToWDateString(DBDATE(strNP09))))
         strNP08 = DBDATE(DateAdd("m", -2, ChangeWStringToWDateString(DBDATE(strNP09))))
      End If
      strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      '下一程序序號
      strNP22 = GetNextProgressNo()
      'Add By Sindy 2009/11/04 判斷是否已掛使用宣誓期限
      bInsert = True
      strSql = "SELECT * FROM NextProgress " & _
               "WHERE NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND " & _
                     "NP06 IS NULL AND " & _
                     "NP07 = '" & strNP07 & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         bInsert = False
      End If
      '2009/11/04 End
      '新增下一程序檔
      '2006/3/10 MODIFY BY SONIA NP10應為智權人員不可為操作人員
      'strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
      '         "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "',105," & _
      '                     strNP08 & "," & strNP09 & ",'" & strUserNum & "'," & strNP22 & ")"
      'Modify By Sindy 2009/11/04
      If bInsert = False Then
         strSql = "Update NextProgress Set NP01='" & strCP09 & "',NP08=" & strNP08 & ", NP09=" & strNP09 & " " & _
                   "WHERE NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND " & _
                     "NP06 IS NULL AND " & _
                     "NP07 = '" & strNP07 & "' "
      '2009/11/04 End
      Else
         '2011/12/15 modify by sonia 智權人員改抓PUB_GetAKindSalesNo
         'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "',105," & _
                              strNP08 & "," & strNP09 & ",'" & m_CP13 & "'," & strNP22 & ")"
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','" & strNP07 & "'," & _
                              strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
      '2006/3/10 END
      End If
      cnnConnection.Execute strSql
      
      'Add By Sindy 2018/1/25 波多黎各112:若申請進度的備註裡有未提使用宣誓,再加掛使用宣誓期限
      If m_TM10 = "112" And m_CP10 = "101" And InStr(m_CP64, "未提使用宣誓") > 0 Then
         '法定期限
         strNP09 = DBDATE(DateAdd("yyyy", 6, ChangeWStringToWDateString(DBDATE(textTM11.Text))))
         '本所期限
         strNP08 = DBDATE(DateAdd("m", -2, ChangeWStringToWDateString(DBDATE(strNP09))))
         strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
         '下一程序序號
         strNP22 = GetNextProgressNo()
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "',105," & _
                              strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
         cnnConnection.Execute strSql
      End If
      '2018/1/25 END
   End If
   
   'Add by Sindy 2023/4/24
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm030102_01", strCP09
   End If
   '2023/4/24 END
   
'Move By Cheng 2002/11/29
'911106 nick transation
cnnConnection.CommitTrans

   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 列印定稿
   If textPrint <> "N" Then
      PrintLetter
   End If
'Modify By Cheng 2002/11/29
''911106 nick transation
'cnnConnection.CommitTrans
Exit Function
CheckingErr:
    MsgBox (Err.Description)
     cnnConnection.RollbackTrans
     OnSaveData = False
End Function

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse

   CheckDataValid = False

   ' 申請日
   If IsEmptyText(textTM11) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入申請日"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      'Add By Cheng 2002/12/17
      Me.textTM11.SetFocus
      textTM11_GotFocus
      GoTo EXITSUB
   End If
    'Add By Cheng 2002/12/17
   ' 回音
    'Modify By Cheng 2003/01/01
    '若有輸入申請案號
    If Me.textTM12.Text <> "" Then
        If IsEmptyText(Me.textCF09.Text) = True Then
           strTit = "檢核資料"
           strMsg = "請輸入回音"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           Me.textCF09.SetFocus
           textCF09_GotFocus
           GoTo EXITSUB
        End If
    End If
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 將所有的文字反白
Private Sub InverseAll(ByRef tb As TextBox)
   tb.SelStart = 0
   tb.SelLength = Len(tb.Text)
End Sub

Private Sub textTM11_GotFocus()
   InverseAll textTM11
End Sub
' 申請日
Private Sub textTM11_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM11) = False Then
      If CheckIsDate(textTM11, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請日日期不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
      'add by sonia 2018/2/6
      If Val(DBDATE(textTM11)) < Val(m_CP27) And Cancel = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "申請日不可小於發文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
      'end 2018/2/6
   End If
   
   If Cancel Then TextInverse textTM11
End Sub

Private Sub textTM12_GotFocus()
   InverseAll textTM12
End Sub

Private Sub textPrint_GotFocus()
   InverseAll textPrint
End Sub

Private Sub textCF09_GotFocus()
'   InverseAll textCF09
    Me.textCF09.SelStart = 0
    Me.textCF09.SelLength = 0
End Sub

Private Sub textCP45_GotFocus()
   InverseAll textCP45
End Sub

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
Dim strSql As String
'Add By Cheng 2003/01/01
Dim strAttach As String '附件
Dim ii As Integer
   
   'add by nickc 2006/10/20
   Select Case m_TM01
   Case "CFT"
      'Modify By Cheng 2003/01/01
      '案號有收到(有輸入案號)
      If Me.textTM12.Text <> "" Then
         ' 清除定稿例外欄位檔原有資料
         EndLetter "02", m_CP09, "01", strUserNum
         ' 回音
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & "02" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & _
                  "','回音','" & textCF09 & "')"
         cnnConnection.Execute strSql
         'Add By Cheng 2003/01/01
         ' 附件
         strAttach = ""
         For ii = 0 To Me.Check1.Count - 1
             If Me.Check1(ii).Value = vbChecked Then strAttach = strAttach & Me.Check1(ii).Caption & "、"
         Next ii
         If strAttach <> "" Then
             strAttach = Left(strAttach, Len(strAttach) - 1)
             strAttach = "附件：" & strAttach & "。"
             strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                      "VALUES ('" & "02" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & _
                      "','附件','" & strAttach & "')"
             cnnConnection.Execute strSql
         End If
      '案號未收到(未輸入案號)
      Else
         ' 清除定稿例外欄位檔原有資料
         EndLetter "02", m_CP09, "02", strUserNum
          'Add By Cheng 2003/01/01
          ' 附件
          strAttach = ""
          For ii = 0 To Me.Check1.Count - 1
              If Me.Check1(ii).Value = vbChecked Then strAttach = strAttach & Me.Check1(ii).Caption & "、"
          Next ii
          If strAttach <> "" Then
              strAttach = Left(strAttach, Len(strAttach) - 1)
              strAttach = "附件：" & strAttach & "。"
              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                       "VALUES ('" & "02" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & _
                       "','附件','" & strAttach & "')"
              cnnConnection.Execute strSql
          End If
      End If
   Case "CFC"
      ' 清除定稿例外欄位檔原有資料
      EndLetter "02", m_CP09, "01", strUserNum
      ' 回音
      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & "02" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & _
               "','回音','" & textCF09 & "')"
      cnnConnection.Execute strSql
   Case Else
   End Select
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   
   'add by nickc 2006/10/20
   Select Case m_TM01
   Case "CFT"
   ' 列印定稿
    'Modify By Cheng 2003/01/01
'   NowPrint m_CP09, "02", "01", False, strUserNum, 0
    '若有輸入申請案號
    If Me.textTM12.Text <> "" Then
        NowPrint m_CP09, "02", "01", False, strUserNum, 0
    '若未輸入申請案號
    Else
        NowPrint m_CP09, "02", "02", False, strUserNum, 0
    End If
   Case "CFC"
        NowPrint m_CP09, "02", "01", False, strUserNum, 0
   Case Else
   End Select
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False

'2011/4/27 cancel by sonia CFT案件不必檢查
''Add By Sindy 2010/12/24
'If Me.textTM12.Enabled = True Then
'   Cancel = False
'   textTM12_Validate Cancel
'   If Cancel = True Then
'      textTM12.SetFocus
'      Exit Function
'   End If
'End If

If Me.textPrint.Enabled = True Then
   Cancel = False
   textPrint_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textTM11.Enabled = True Then
   Cancel = False
   textTM11_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Added by Lydia 2021/08/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
     Exit Function
End If
   
TxtValidate = True
End Function

'2011/4/27 cancel by sonia CFT案件不必檢查
''Add By Sindy 2010/9/1
'Private Sub textTM12_Validate(Cancel As Boolean)
'   If IsEmptyText(textTM12) = False Then
'      '檢查申請案號所輸入的長度是否正確
'      If PUB_ChkTm12Tm15Length("1", textTM12, m_TM01, m_TM02, m_TM03, m_TM04, m_TM10) = False Then
'         Cancel = True
'         textTM12_GotFocus
'         Exit Sub
'      End If
'   End If
'End Sub
